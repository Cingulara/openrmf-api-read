using System;
using System.Collections.Generic;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Moq;
using NATS.Client;
using openrmf_read_api.Controllers;
using openrmf_read_api.Data;
using openrmf_read_api.Models;
using Xunit;

namespace tests.Controllers;

public class ControllerBehaviorTests
{
    [Fact]
    public void HealthController_Get_ReturnsOk_WhenRepositoryIsHealthy()
    {
        var repo = new Mock<ISystemGroupRepository>();
        repo.Setup(r => r.HealthStatus()).Returns(true);
        var sut = new HealthController(repo.Object, MockLogger<HealthController>());

        var result = sut.Get();

        var ok = Assert.IsType<OkObjectResult>(result.Result);
        Assert.Equal("ok", ok.Value);
        Assert.NotEqual("database error", ok.Value);
    }

    [Fact]
    public void HealthController_Get_ReturnsBadRequest_WhenRepositoryIsUnhealthy()
    {
        var repo = new Mock<ISystemGroupRepository>();
        repo.Setup(r => r.HealthStatus()).Returns(false);
        var sut = new HealthController(repo.Object, MockLogger<HealthController>());

        var result = sut.Get();

        var bad = Assert.IsType<BadRequestObjectResult>(result.Result);
        Assert.Equal("database error", bad.Value);
        Assert.NotEqual("ok", bad.Value);
    }

    [Fact]
    public void HealthController_Get_ReturnsBadRequest_WhenRepositoryThrows()
    {
        var repo = new Mock<ISystemGroupRepository>();
        repo.Setup(r => r.HealthStatus()).Throws(new InvalidOperationException("boom"));
        var sut = new HealthController(repo.Object, MockLogger<HealthController>());

        var result = sut.Get();

        var bad = Assert.IsType<BadRequestObjectResult>(result.Result);
        Assert.Equal("Improper API configuration", bad.Value);
        Assert.NotEqual("ok", bad.Value);
    }

    [Fact]
    public async Task ReadController_ListArtifactSystems_ReturnsOk_AndSanitizesRawPatchData()
    {
        var artifactRepo = new Mock<IArtifactRepository>();
        var systemRepo = new Mock<ISystemGroupRepository>();
        systemRepo
            .Setup(r => r.GetAllSystemGroups())
            .ReturnsAsync(new List<SystemGroup>
            {
                new SystemGroup { title = "System A", rawNessusFile = "raw-data" }
            });

        var sut = new ReadController(artifactRepo.Object, systemRepo.Object, MockLogger<ReadController>());

        var result = await sut.ListArtifactSystems();

        var ok = Assert.IsType<OkObjectResult>(result);
        var systems = Assert.IsAssignableFrom<IEnumerable<SystemGroup>>(ok.Value);
        Assert.NotEmpty(systems);
        Assert.Equal(string.Empty, Assert.Single(systems).rawNessusFile);
    }

    [Fact]
    public async Task ReadController_ListArtifactSystems_ReturnsNotFound_WhenRepositoryReturnsNull()
    {
        var artifactRepo = new Mock<IArtifactRepository>();
        var systemRepo = new Mock<ISystemGroupRepository>();
        systemRepo.Setup(r => r.GetAllSystemGroups()).ReturnsAsync((IEnumerable<SystemGroup>)null!);

        var sut = new ReadController(artifactRepo.Object, systemRepo.Object, MockLogger<ReadController>());

        var result = await sut.ListArtifactSystems();

        Assert.IsType<NotFoundResult>(result);
    }

    [Fact]
    public async Task ReadController_ExportChecklistListing_ReturnsBadRequest_WhenSystemIsMissing()
    {
        var artifactRepo = new Mock<IArtifactRepository>();
        var systemRepo = new Mock<ISystemGroupRepository>();
        var sut = new ReadController(artifactRepo.Object, systemRepo.Object, MockLogger<ReadController>());

        var result = await sut.ExportChecklistListing(string.Empty);

        var bad = Assert.IsType<BadRequestObjectResult>(result);
        Assert.Equal("You must specify a system to export", bad.Value);
        Assert.NotEqual("ok", bad.Value);
        artifactRepo.Verify(r => r.GetSystemArtifacts(It.IsAny<string>()), Times.Never);
    }

    [Fact]
    public async Task ComplianceController_GetComplianceBySystem_ReturnsBadRequest_WhenIdMissing()
    {
        var artifactRepo = new Mock<IArtifactRepository>();
        var systemRepo = new Mock<ISystemGroupRepository>();
        var sut = new ComplianceController(artifactRepo.Object, systemRepo.Object, MockLogger<ComplianceController>());

        var result = await sut.GetCompliancBySystem(string.Empty, "low", false);

        Assert.IsType<BadRequestResult>(result);
    }

    [Fact]
    public async Task ComplianceController_GetComplianceBySystem_ReturnsBadRequest_WhenRepositoryThrows()
    {
        var artifactRepo = new Mock<IArtifactRepository>();
        artifactRepo
            .Setup(r => r.GetSystemArtifacts("sys1"))
            .ThrowsAsync(new InvalidOperationException("db down"));
        var systemRepo = new Mock<ISystemGroupRepository>();
        var sut = new ComplianceController(artifactRepo.Object, systemRepo.Object, MockLogger<ComplianceController>());

        var result = await sut.GetCompliancBySystem("sys1", "low", false);

        Assert.IsType<BadRequestResult>(result);
    }

    [Fact]
    public async Task ComplianceController_GetComplianceBySystemExport_ReturnsNotFound_WhenSystemMissing()
    {
        var artifactRepo = new Mock<IArtifactRepository>();
        var systemRepo = new Mock<ISystemGroupRepository>();
        systemRepo.Setup(r => r.GetSystemGroup("sys1")).ReturnsAsync((SystemGroup)null!);
        var sut = new ComplianceController(artifactRepo.Object, systemRepo.Object, MockLogger<ComplianceController>());

        var result = await sut.GetCompliancBySystemExport("sys1", "low", false);

        Assert.IsType<NotFoundResult>(result);
    }

    [Fact]
    public async Task UploadController_UploadNewChecklist_ReturnsBadRequest_WhenNoFilesProvided()
    {
        var artifactRepo = new Mock<IArtifactRepository>();
        var systemRepo = new Mock<ISystemGroupRepository>();
        var msgConnection = new Mock<IConnection>();
        var options = Microsoft.Extensions.Options.Options.Create(new NATSServer { connection = msgConnection.Object });
        var sut = new UploadController(artifactRepo.Object, MockLogger<UploadController>(), options, systemRepo.Object);

        var result = await sut.UploadNewChecklist(new List<IFormFile>(), string.Empty, "my-system");

        Assert.IsType<BadRequestResult>(result);
    }

    [Fact]
    public async Task UploadController_CreateNewChecklist_ReturnsNotFound_WhenSystemMissing()
    {
        var artifactRepo = new Mock<IArtifactRepository>();
        var systemRepo = new Mock<ISystemGroupRepository>();
        systemRepo.Setup(r => r.GetSystemGroup("sys-404")).ReturnsAsync((SystemGroup)null!);

        var msgConnection = new Mock<IConnection>();
        var options = Microsoft.Extensions.Options.Options.Create(new NATSServer { connection = msgConnection.Object });
        var sut = new UploadController(artifactRepo.Object, MockLogger<UploadController>(), options, systemRepo.Object);
        AttachUserClaim(sut);

        var result = await sut.CreateNewChecklist("sys-404", "template-1");

        Assert.IsType<NotFoundResult>(result);
        artifactRepo.Verify(r => r.AddArtifact(It.IsAny<Artifact>()), Times.Never);
    }

    [Fact]
    public async Task SaveController_DeleteArtifact_ReturnsNotFound_WhenArtifactMissing()
    {
        var artifactRepo = new Mock<IArtifactRepository>();
        artifactRepo.Setup(r => r.GetArtifact("a1")).ReturnsAsync((Artifact)null!);

        var systemRepo = new Mock<ISystemGroupRepository>();
        var msgConnection = new Mock<IConnection>();
        var options = Microsoft.Extensions.Options.Options.Create(new NATSServer { connection = msgConnection.Object });
        var sut = new SaveController(artifactRepo.Object, systemRepo.Object, MockLogger<SaveController>(), options);

        var result = await sut.DeleteArtifact("a1");

        Assert.IsType<NotFoundResult>(result);
        msgConnection.Verify(c => c.Publish(It.IsAny<string>(), It.IsAny<byte[]>()), Times.Never);
    }

    [Fact]
    public async Task SaveController_DeleteArtifact_ReturnsBadRequest_WhenRepositoryThrows()
    {
        var artifactRepo = new Mock<IArtifactRepository>();
        artifactRepo.Setup(r => r.GetArtifact("a1")).ThrowsAsync(new InvalidOperationException("db failure"));

        var systemRepo = new Mock<ISystemGroupRepository>();
        var msgConnection = new Mock<IConnection>();
        var options = Microsoft.Extensions.Options.Options.Create(new NATSServer { connection = msgConnection.Object });
        var sut = new SaveController(artifactRepo.Object, systemRepo.Object, MockLogger<SaveController>(), options);

        var result = await sut.DeleteArtifact("a1");

        Assert.IsType<BadRequestResult>(result);
    }

    [Fact]
    public async Task SaveController_UpdateSystem_ReturnsNotFound_WhenSystemMissing()
    {
        var artifactRepo = new Mock<IArtifactRepository>();
        var systemRepo = new Mock<ISystemGroupRepository>();
        systemRepo.Setup(r => r.GetSystemGroup("sys-none")).ReturnsAsync((SystemGroup)null!);

        var msgConnection = new Mock<IConnection>();
        var options = Microsoft.Extensions.Options.Options.Create(new NATSServer { connection = msgConnection.Object });
        var sut = new SaveController(artifactRepo.Object, systemRepo.Object, MockLogger<SaveController>(), options);

        var result = await sut.UpdateSystem("sys-none", "title", "desc", null!);

        Assert.IsType<NotFoundResult>(result);
        msgConnection.Verify(c => c.Publish(It.IsAny<string>(), It.IsAny<byte[]>()), Times.Never);
    }

    [Fact]
    public async Task SaveController_DeleteSystemPatchScanFile_ReturnsNotFound_WhenSystemMissing()
    {
        var artifactRepo = new Mock<IArtifactRepository>();
        var systemRepo = new Mock<ISystemGroupRepository>();
        systemRepo.Setup(r => r.GetSystemGroup("sys-none")).ReturnsAsync((SystemGroup)null!);

        var msgConnection = new Mock<IConnection>();
        var options = Microsoft.Extensions.Options.Options.Create(new NATSServer { connection = msgConnection.Object });
        var sut = new SaveController(artifactRepo.Object, systemRepo.Object, MockLogger<SaveController>(), options);

        var result = await sut.DeleteSystemPatchScanFile("sys-none");

        Assert.IsType<NotFoundResult>(result);
        msgConnection.Verify(c => c.Publish(It.IsAny<string>(), It.IsAny<byte[]>()), Times.Never);
    }

    private static ILogger<T> MockLogger<T>()
    {
        return new Mock<ILogger<T>>().Object;
    }

    private static void AttachUserClaim(Controller controller)
    {
        var identity = new ClaimsIdentity(
        [
            new Claim(ClaimTypes.NameIdentifier, Guid.NewGuid().ToString()),
            new Claim("name", "Unit Tester"),
            new Claim("preferred_username", "unit.tester"),
            new Claim("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress", "tester@example.com")
        ],
        "test");

        controller.ControllerContext = new ControllerContext
        {
            HttpContext = new DefaultHttpContext
            {
                User = new ClaimsPrincipal(identity)
            }
        };
    }
}
