using System;
using System.Collections.Generic;
using System.Reflection;
using openrmf_read_api.Classes;
using openrmf_read_api.Models;
using Xunit;

namespace tests.Classes;

public class UtilityClassTests
{
    [Fact]
    public void Compression_CompressAndDecompress_RoundTripsData()
    {
        var original = "OpenRMF test payload";

        var compressed = Compression.CompressString(original);
        var restored = Compression.DecompressString(compressed);

        Assert.Equal(original, restored);
        Assert.NotEqual("different", restored);
    }

    [Fact]
    public void Compression_Decompress_ThrowsOnInvalidInput()
    {
        Assert.Throws<FormatException>(() => Compression.DecompressString("not-base64"));
    }

    [Fact]
    public void RecordGenerator_CleanData_RemovesTabsAndLineBreakSeparators()
    {
        const string dirty = "<a>\n\t</a>\n<z>\n</z>";

        var cleaned = RecordGenerator.CleanData(dirty);

        Assert.DoesNotContain("\t", cleaned);
        Assert.DoesNotContain(">\n<", cleaned);
    }

    [Fact]
    public void RecordGenerator_DecodeHtml_HandlesNullAndEncodedValues()
    {
        var decoded = RecordGenerator.DecodeHTML("A&amp;B");
        var decodedNull = RecordGenerator.DecodeHTML(null!);

        Assert.Equal("A&B", decoded);
        Assert.NotEqual("A&amp;B", decoded);
        Assert.Equal(string.Empty, decodedNull);
    }

    [Fact]
    public void NessusPatchLoader_SanitizeHostname_MasksIpv4AndKeepsHostnames()
    {
        var masked = NessusPatchLoader.SanitizeHostname("10.23.45.67");
        var host = NessusPatchLoader.SanitizeHostname("server-01");

        Assert.Equal("xxx.xxx.45.67", masked);
        Assert.Equal("server-01", host);
        Assert.NotEqual("10.23.45.67", masked);
    }

    [Fact]
    public void NessusPatchLoader_LoadPatchData_ParsesReportAndHostEntries()
    {
        const string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<NessusClientData_v2><Report name=\"Weekly Scan\">"
            + "<ReportHost name=\"10.20.30.40\">"
            + "<HostProperties>"
            + "<tag name=\"hostname\">host-1</tag>"
            + "<tag name=\"operating-system\">Linux</tag>"
            + "<tag name=\"host-ip\">10.20.30.40</tag>"
            + "<tag name=\"HOST_END_TIMESTAMP\">1700000000</tag>"
            + "</HostProperties>"
            + "<ReportItem severity=\"2\" pluginID=\"100\" pluginName=\"Plugin\" pluginFamily=\"Family\" port=\"0\" svc_name=\"general\" protocol=\"tcp\">"
            + "<description>desc</description><risk_factor>Medium</risk_factor>"
            + "</ReportItem></ReportHost></Report></NessusClientData_v2>";

        var result = NessusPatchLoader.LoadPatchData(xml);

        Assert.Equal("Weekly Scan", result.reportName);
        Assert.Single(result.summary);
        Assert.NotEqual(string.Empty, result.summary[0].hostname);
    }

    [Fact]
    public void ChecklistLoader_LoadChecklist_ParsesAssetStigAndVulnerability()
    {
        const string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<CHECKLIST><ASSET><ROLE>None</ROLE><ASSET_TYPE>Computing</ASSET_TYPE><MARKING>U</MARKING>"
            + "<HOST_NAME>host-1</HOST_NAME><HOST_IP>1.1.1.1</HOST_IP><HOST_MAC>AA</HOST_MAC>"
            + "<HOST_FQDN>host-1.local</HOST_FQDN><TECH_AREA>Tech</TECH_AREA><TARGET_KEY>key</TARGET_KEY>"
            + "<WEB_OR_DATABASE>false</WEB_OR_DATABASE><WEB_DB_SITE></WEB_DB_SITE><WEB_DB_INSTANCE></WEB_DB_INSTANCE></ASSET>"
            + "<STIGS><iSTIG><STIG_INFO><SI_DATA><SID_NAME>title</SID_NAME><SID_DATA>My STIG</SID_DATA></SI_DATA></STIG_INFO>"
            + "<VULN><STIG_DATA><VULN_ATTRIBUTE>Vuln_Num</VULN_ATTRIBUTE><ATTRIBUTE_DATA>V-1</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Rule_Ver</VULN_ATTRIBUTE><ATTRIBUTE_DATA>SV-1r1_rule</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>CCI_REF</VULN_ATTRIBUTE><ATTRIBUTE_DATA>CCI-000001</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STATUS>Open</STATUS><FINDING_DETAILS>details</FINDING_DETAILS><COMMENTS>comments</COMMENTS>"
            + "<SEVERITY_OVERRIDE></SEVERITY_OVERRIDE><SEVERITY_JUSTIFICATION></SEVERITY_JUSTIFICATION>"
            + "</VULN></iSTIG></STIGS></CHECKLIST>";

        var checklist = ChecklistLoader.LoadChecklist(xml);

        Assert.Equal("host-1", checklist.ASSET.HOST_NAME);
        Assert.Single(checklist.STIGS.iSTIG.VULN);
        Assert.NotEqual(string.Empty, checklist.STIGS.iSTIG.VULN[0].STATUS);
    }

    [Fact]
    public void ChecklistLoader_UpdateChecklistVulnerabilityOrder_ReturnsNormalizedXml()
    {
        const string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<CHECKLIST><ASSET><ROLE>None</ROLE><ASSET_TYPE>Computing</ASSET_TYPE><MARKING>U</MARKING>"
            + "<HOST_NAME>host-1</HOST_NAME><HOST_IP>1.1.1.1</HOST_IP><HOST_MAC>AA</HOST_MAC>"
            + "<HOST_FQDN>host-1.local</HOST_FQDN><TECH_AREA>Tech</TECH_AREA><TARGET_KEY>key</TARGET_KEY>"
            + "<WEB_OR_DATABASE>false</WEB_OR_DATABASE><WEB_DB_SITE></WEB_DB_SITE><WEB_DB_INSTANCE></WEB_DB_INSTANCE></ASSET>"
            + "<STIGS><iSTIG><STIG_INFO><SI_DATA><SID_NAME>title</SID_NAME><SID_DATA>My STIG</SID_DATA></SI_DATA></STIG_INFO>"
            + "<VULN>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Vuln_Num</VULN_ATTRIBUTE><ATTRIBUTE_DATA>V-1</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Severity</VULN_ATTRIBUTE><ATTRIBUTE_DATA>high</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Group_Title</VULN_ATTRIBUTE><ATTRIBUTE_DATA>grp</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Rule_ID</VULN_ATTRIBUTE><ATTRIBUTE_DATA>R-1</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Rule_Ver</VULN_ATTRIBUTE><ATTRIBUTE_DATA>SV-1r1_rule</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Rule_Title</VULN_ATTRIBUTE><ATTRIBUTE_DATA>title</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Vuln_Discuss</VULN_ATTRIBUTE><ATTRIBUTE_DATA>d</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>IA_Controls</VULN_ATTRIBUTE><ATTRIBUTE_DATA>ia</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Check_Content</VULN_ATTRIBUTE><ATTRIBUTE_DATA>c</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Fix_Text</VULN_ATTRIBUTE><ATTRIBUTE_DATA>f</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>False_Positives</VULN_ATTRIBUTE><ATTRIBUTE_DATA></ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>False_Negatives</VULN_ATTRIBUTE><ATTRIBUTE_DATA></ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Documentable</VULN_ATTRIBUTE><ATTRIBUTE_DATA></ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Mitigations</VULN_ATTRIBUTE><ATTRIBUTE_DATA></ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Potential_Impact</VULN_ATTRIBUTE><ATTRIBUTE_DATA></ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Third_Party_Tools</VULN_ATTRIBUTE><ATTRIBUTE_DATA></ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Mitigation_Control</VULN_ATTRIBUTE><ATTRIBUTE_DATA></ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Responsibility</VULN_ATTRIBUTE><ATTRIBUTE_DATA></ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Security_Override_Guidance</VULN_ATTRIBUTE><ATTRIBUTE_DATA></ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Check_Content_Ref</VULN_ATTRIBUTE><ATTRIBUTE_DATA></ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Weight</VULN_ATTRIBUTE><ATTRIBUTE_DATA>10</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>Class</VULN_ATTRIBUTE><ATTRIBUTE_DATA>Unclass</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>STIGRef</VULN_ATTRIBUTE><ATTRIBUTE_DATA>ref</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>TargetKey</VULN_ATTRIBUTE><ATTRIBUTE_DATA>tk</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>STIG_UUID</VULN_ATTRIBUTE><ATTRIBUTE_DATA>uuid</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STIG_DATA><VULN_ATTRIBUTE>CCI_REF</VULN_ATTRIBUTE><ATTRIBUTE_DATA>CCI-000001</ATTRIBUTE_DATA></STIG_DATA>"
            + "<STATUS>Open</STATUS><FINDING_DETAILS></FINDING_DETAILS><COMMENTS></COMMENTS>"
            + "<SEVERITY_OVERRIDE></SEVERITY_OVERRIDE><SEVERITY_JUSTIFICATION></SEVERITY_JUSTIFICATION>"
            + "</VULN></iSTIG></STIGS></CHECKLIST>";

        var normalized = ChecklistLoader.UpdateChecklistVulnerabilityOrder(xml);

        Assert.Contains("<CHECKLIST>", normalized);
        Assert.Contains("Vuln_Num", normalized);
        Assert.DoesNotContain("\t", normalized);
    }

    [Fact]
    public void ScapScanResultLoader_LoadScapScan_ParsesTitleAndRuleResults()
    {
        const string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + "<cdf:Benchmark xmlns:cdf=\"http://checklists.nist.gov/xccdf/1.2\">"
            + "<cdf:title>My STIG Title</cdf:title>"
            + "<cdf:target>host-01</cdf:target>"
            + "<cdf:TestResult end-time=\"2026-05-15T12:00:00\" test-system=\"SCC\">"
            + "<cdf:rule-result idref=\"xccdf_mil.disa.stig_rule_SV-1r1_rule\" version=\"SV-1r1_rule\">"
            + "<cdf:result>pass</cdf:result>"
            + "</cdf:rule-result>"
            + "</cdf:TestResult>"
            + "</cdf:Benchmark>";

        var result = SCAPScanResultLoader.LoadSCAPScan(xml);

        Assert.Equal("My STIG Title", result.title);
        Assert.Single(result.ruleResults);
        Assert.Equal("pass", result.ruleResults[0].result);
        Assert.NotEqual("fail", result.ruleResults[0].result);
    }

    [Fact]
    public void ComplianceGenerator_PrivateMethods_ProduceExpectedOutputs()
    {
        var generatorType = typeof(ComplianceGenerator);

        var getFirstIndex = generatorType.GetMethod("GetFirstIndex", BindingFlags.NonPublic | BindingFlags.Static)!;
        var generateStatus = generatorType.GetMethod("GenerateStatus", BindingFlags.NonPublic | BindingFlags.Static)!;
        var generateSort = generatorType.GetMethod("GenerateControlIndexSort", BindingFlags.NonPublic | BindingFlags.Static)!;
        var createList = generatorType.GetMethod("CreateListOfNISTControls", BindingFlags.NonPublic | BindingFlags.Static)!;

        var firstIndex = (int)getFirstIndex.Invoke(null, new object[] { "AC-2 (1)" })!;
        var status = (string)generateStatus.Invoke(null, new object[] { "notafinding", "open" })!;
        var sort = (string)generateSort.Invoke(null, new object[] { "AC-2.3" })!;

        var cci = new CciItem
        {
            cciId = "CCI-000001",
            references = new List<CciReference>
            {
                new CciReference
                {
                    majorControl = "AC-2",
                    index = "AC-2.3",
                    location = "NIST",
                    title = "Account Management",
                    version = "rev5"
                }
            }
        };

        var controls = (List<openrmf_read_api.Models.Compliance.NISTControl>)createList.Invoke(null, new object[] { new List<CciItem> { cci } })!;

        Assert.Equal(4, firstIndex);
        Assert.NotEqual(-1, firstIndex);
        Assert.Equal("open", status);
        Assert.NotEqual("notafinding", status);
        Assert.Equal("AC-02", sort);
        Assert.Single(controls);
        Assert.Equal("CCI-000001", controls[0].CCI);
    }
}
