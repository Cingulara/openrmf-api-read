
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using openstig_read_api.Classes;
using openstig_read_api.Models;
using System.IO;
using System.Text;
using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Hosting;
using System.Xml.Serialization;
using System.Xml;
using Microsoft.AspNetCore.Cors.Infrastructure;
using Newtonsoft.Json;
using Microsoft.Extensions.Options;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Distributed;
using Microsoft.Extensions.Logging;

using openstig_read_api.Data;

namespace openstig_read_api.Controllers
{
    [Route("[controller]")]
    public class ReadController : Controller
    {
	    private readonly IArtifactRepository _artifactRepo;
        private readonly ILogger<ReadController> _logger;

        public ReadController(IArtifactRepository artifactRepo, ILogger<ReadController> logger)
        {
            _logger = logger;
            _artifactRepo = artifactRepo;
        }

        // GET values
        [HttpGet]
        public async Task<IActionResult> ListArtifacts()
        {
            try {
                IEnumerable<Artifact> artifacts;
                artifacts = await _artifactRepo.GetAllArtifacts();
                foreach (Artifact a in artifacts) {
                    // deserialize the checklist so it comes out more gooder
                    a.Checklist = ChecklistLoader.LoadChecklist(a.rawChecklist);
                    a.rawChecklist = string.Empty;
                }
                return Ok(artifacts);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error listing all artifacts and deserializing the checklist XML");
                return BadRequest();
            }
        }

        // GET /value
        [HttpGet("{id}")]
        public async Task<IActionResult> GetArtifacts(string id)
        {
            try {
                Artifact art = new Artifact();
                art = await _artifactRepo.GetArtifact(id);
                art.Checklist = ChecklistLoader.LoadChecklist(art.rawChecklist);
                //art.rawChecklist = string.Empty;
                return Ok(art);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error Retrieving Artifact");
                return NotFound();
            }
        }
        
    }
}
