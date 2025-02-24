namespace Tap.Zis3.Domain.Services.Reports.TemplateDtos;

using System;
using System.Collections.Generic;

public struct InformationAnalyticCardForComplexTemplateDto
{
    public List<LandPlotProperties> LandPlotProperties { get; set; }
    public CommonComplexProperties CommonComplexProperties { get; set; }
}

public struct LandPlotProperties
{
    public string Area { get; set; }
    public string LandPlotRightsType { get; set; }
    public string plotNotFormed { get; set; }
    public string plotFormed { get; set; }
}

public struct CommonComplexProperties
{
    public string AreaMeassure { get; set; }
    public string TotalBuildingsArea { get; set; }
    public string Address { get; set; }
}
