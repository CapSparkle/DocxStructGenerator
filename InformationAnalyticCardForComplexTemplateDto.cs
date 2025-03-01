using System;
using System.Collections.Generic;

public class InformationAnalyticCardForComplexTemplateDto
{
    public string Owner { get; set; }
    public string NameComplexObject { get; set; }
    public string CreateDocumentDate { get; set; }
    public string ConclusionsFromPreliminaryAnalysis { get; set; }
    public string ProposedCourseOfAction { get; set; }
    public string JustificationActionAndImplementation { get; set; }
    public string FinancialEffectModel { get; set; }
    public string SalesActivitiesCarriedOut { get; set; }
    public string OtherSalesActivities { get; set; }
    public string ActionPlan { get; set; }
    public string Problems { get; set; }
    public string AreaOfPremisesLeased { get; set; }
    public List<LandPlotProperties> LandPlotProperties { get; set; } = new List<LandPlotProperties>();
    public List<CompositionOfTheComplex> CompositionOfTheComplex { get; set; } = new List<CompositionOfTheComplex>();
    public List<CompetitiveProcedures> CompetitiveProceduresList0 { get; set; } = new List<CompetitiveProcedures>();
    public List<TechnologicalComplex> TechnologicalComplexList0 { get; set; } = new List<TechnologicalComplex>();
    public List<AlienationOfAssets> AlienationOfAssetsList0 { get; set; } = new List<AlienationOfAssets>();
    public CommonComplexProperties CommonComplexProperties { get; set; }
    public IncomeAndExpenses IncomeAndExpenses { get; set; }
}

public class CompositionOfTheComplex
{
    public string blockIndex { get; set; }
    public string IsBuilding { get; set; }
    public string IsConstraction { get; set; }
    public string IsUnfinishedConstruction { get; set; }
    public string IsMovables { get; set; }
    public string IsAnother { get; set; }
    public string ObjectType { get; set; }
    public string ObjectNameOfEgrn { get; set; }
    public string ObjectNameOfAccounting { get; set; }
    public CommonObjectProperties CommonObjectProperties { get; set; }
}

public class CommonObjectProperties
{
    public string Address { get; set; }
    public string ConstructionYear { get; set; }
    public string Purpose { get; set; }
    public string CadastralNumber { get; set; }
    public string InventoryNumber { get; set; }
    public string InventoryBookkeepingNumber { get; set; }
    public string CertificateOfOwnership { get; set; }
    public string DetailsRegister { get; set; }
    public string Restrictions { get; set; }
    public string MainCharacteristicsType { get; set; }
    public string MainCharacteristicsValue { get; set; }
    public string MainCharacteristicsMeasurementUnit { get; set; }
    public string FloorCount { get; set; }
    public string UnrergroundFloorCount { get; set; }
    public string DegreeOfReadiness { get; set; }
    public string YearOfCompletion { get; set; }
    public string NameNetworks { get; set; }
    public string PowerSupply { get; set; }
    public string WaterSuply { get; set; }
    public string HeatSuply { get; set; }
    public string TransportAccessibility { get; set; }
    public string LocationToRestrictedAreas { get; set; }
    public string CurrentUse { get; set; }
    public string PossibleUse { get; set; }
    public string ReportTaxDate { get; set; }
    public string StartTaxCost { get; set; }
    public string ResidualTaxCost { get; set; }
    public string ReportBookkeepingDate { get; set; }
    public string StartBookkeepingCost { get; set; }
    public string ResidualBookkeepingCost { get; set; }
    public string ReportMarketDate { get; set; }
    public string MarketCostWithDuty { get; set; }
    public string MarketCostDutyFree { get; set; }
    public string NameOfAppraiser { get; set; }
}

public class IncomeAndExpenses
{
    public string PrevYear { get; set; }
    public string СurrentYear { get; set; }
    public string ActualTotal { get; set; }
    public string СurrentYearTotal { get; set; }
    public string NotesTotal { get; set; }
    public List<Expenses> ExpensesList0 { get; set; } = new List<Expenses>();
    public List<Income> IncomeList0 { get; set; } = new List<Income>();
}

public class LandPlotProperties
{
    public string TypeOfLandRights { get; set; }
    public string Square { get; set; }
    public string CadastralNumber { get; set; }
    public string LandCategory { get; set; }
    public string CadastralValue { get; set; }
    public string Encumbrances { get; set; }
    public string PermittedUse { get; set; }
    public string LandPaymentsPerYear { get; set; }
    public string Rent { get; set; }
    public string RentalPeriod { get; set; }
    public string BookValue { get; set; }
    public string MarketValue { get; set; }
    public string NameOfAppraiser { get; set; }
    public string PlotFormed { get; set; }
    public string PlotNotFormed { get; set; }
}

public class CommonComplexProperties
{
    public string NameOfAppraiser { get; set; }
    public string ComplexReportTaxDate { get; set; }
    public string StartTaxCost { get; set; }
    public string ResidualTaxCost { get; set; }
    public string ComplexReportBookkeepingDate { get; set; }
    public string StartBookkeepingCost { get; set; }
    public string ResidualBookkeepingCost { get; set; }
    public string ComplexReportMarketDate { get; set; }
    public string MarketValueWithDuty { get; set; }
    public string MarketValueWDutyFree { get; set; }
    public string RestrictionsOfRight { get; set; }
    public string MainCharacteristicTypes { get; set; }
    public string MainCharacteristicValues { get; set; }
    public string MainCharacteristicUnits { get; set; }
    public string NameNetworks { get; set; }
    public string PowerSupply { get; set; }
    public string WaterSupply { get; set; }
    public string HeatSupply { get; set; }
    public string TransportAccessibility { get; set; }
    public string CurrentUsage { get; set; }
    public string PossibleUsage { get; set; }
    public string TotalBuildingsArea { get; set; }
    public string AreaMeassure { get; set; }
    public string Adress { get; set; }
}

public class Income
{
    public string Index { get; set; }
    public string Name { get; set; }
    public string FactPrevYear { get; set; }
    public string FactСurrentYear { get; set; }
    public string Notes { get; set; }
}

public class Expenses
{
    public string Index { get; set; }
    public string Name { get; set; }
    public string FactPrevYear { get; set; }
    public string FactСurrentYear { get; set; }
    public string Notes { get; set; }
}

public class CompetitiveProcedures
{
    public string ResiduaValue { get; set; }
    public string ProcedureName { get; set; }
    public string DateOfPublicationInMedia { get; set; }
    public string StartDate { get; set; }
    public string ClosingDate { get; set; }
    public string EndDate { get; set; }
    public string StartSellingPrice { get; set; }
    public string MinStartSellingPrice { get; set; }
    public string SalesResult { get; set; }
}

public class TechnologicalComplex
{
    public string LeasedOut { get; set; }
    public string WhenRentedOut { get; set; }
    public string Counterparty { get; set; }
    public string RentActivities { get; set; }
    public string RentActionPlan { get; set; }
}

public class AlienationOfAssets
{
    public string Completed { get; set; }
    public string AlienationValue { get; set; }
    public string Other { get; set; }
    public string SalesIncome { get; set; }
    public string FinancialResult { get; set; }
    public string MethodOfAlienation { get; set; }
    public string ContractDetails { get; set; }
    public string ConterpartyName { get; set; }
    public string WriteOffDate { get; set; }
    public string Notes { get; set; }
}
