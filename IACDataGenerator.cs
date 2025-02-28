using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

    public class IACDataGenerator
    {
        public static InformationAnalyticCardForComplexTemplateDto GenerateExampleData()
        {
            var random = new Random();

            return new InformationAnalyticCardForComplexTemplateDto
            {
                Owner = "John Doe",
                NameComplexObject = "Industrial Complex Alpha",
                CreateDocumentDate = DateTime.Now.ToString("yyyy-MM-dd"),
                ConclusionsFromPreliminaryAnalysis = "Initial review completed.",
                ProposedCourseOfAction = "Recommend further evaluation.",
                JustificationActionAndImplementation = "Due to market conditions...",
                FinancialEffectModel = "Projected financial gains: 10%",
                SalesActivitiesCarriedOut = "Direct negotiations conducted.",
                OtherSalesActivities = "Online listings posted.",
                ActionPlan = "Step 1: Evaluation, Step 2: Bidding",
                Problems = "Legal clearance pending.",
                LandPlotProperties = GenerateList(() => new LandPlotProperties
                {
                    TypeOfLandRights = "Lease",
                    Square = random.Next(100, 5000).ToString() + " sq.m",
                    CadastralNumber = "123-456-789",
                    LandCategory = "Industrial",
                    CadastralValue = "$" + random.Next(10000, 500000),
                    Encumbrances = "None",
                    PermittedUse = "Manufacturing",
                    LandPaymentsPerYear = "$" + random.Next(1000, 50000),
                    Rent = "$" + random.Next(5000, 50000),
                    RentalPeriod = random.Next(1, 10) + " years",
                    BookValue = "$" + random.Next(20000, 500000),
                    MarketValue = "$" + random.Next(25000, 550000),
                    NameOfAppraiser = "XYZ Appraisers",
                    PlotFormed = "Yes",
                    PlotNotFormed = "No"
                }),
                CompositionOfTheComplex = GenerateList(() => new CompositionOfTheComplex
                {
                    IsBuilding = "Yes",
                    IsConstraction = "No",
                    IsUnfinishedConstruction = "No",
                    IsAnother = "No",
                    ObjectType = "Factory",
                    ObjectNameOfEgrn = "Factory Alpha",
                    ObjectNameOfAccounting = "Factory A1",
                    CommonObjectProperties = new CommonObjectProperties
                    {
                        Address = "123 Industrial Zone, City",
                        ConstructionYear = "1995",
                        Purpose = "Manufacturing",
                        CadastralNumber = "987-654-321",
                        InventoryNumber = "INV-5678",
                        InventoryBookkeepingNumber = "BK-5678",
                        CertificateOfOwnership = "Cert-001",
                        DetailsRegister = "Reg-345",
                        Restrictions = "None",
                        MainCharacteristicsType = "Factory",
                        MainCharacteristicsValue = "5000 sq.m",
                        MainCharacteristicsMeasurementUnit = "sq.m",
                        FloorCount = "3",
                        UnrergroundFloorCount = "1",
                        DegreeOfReadiness = "100%",
                        YearOfCompletion = "1995",
                        NameNetworks = "Power, Water, Gas",
                        PowerSupply = "Yes",
                        WaterSuply = "Yes",
                        HeatSuply = "Yes",
                        TransportAccessibility = "Highway access",
                        LocationToRestrictedAreas = "None",
                        CurrentUse = "Active Production",
                        PossibleUse = "Storage, Distribution",
                        ReportTaxDate = "2023-01-01",
                        StartTaxCost = "$500,000",
                        ResidualTaxCost = "$250,000",
                        ReportBookkeepingDate = "2023-01-01",
                        StartBookkeepingCost = "$400,000",
                        ResidualBookkeepingCost = "$200,000",
                        ReportMarketDate = "2023-06-01",
                        MarketCostWithDuty = "$600,000",
                        MarketCostDutyFree = "$580,000",
                        NameOfAppraiser = "XYZ Appraisers"
                    }
                }),
                CompetitiveProceduresList0 = GenerateList(() => new CompetitiveProcedures
                {
                    ResiduaValue = "$" + random.Next(10000, 100000),
                    ProcedureName = "Bidding Round " + random.Next(1, 10),
                    DateOfPublicationInMedia = "2023-06-15",
                    StartDate = "2023-07-01",
                    ClosingDate = "2023-07-15",
                    EndDate = "2023-07-20",
                    StartSellingPrice = "$" + random.Next(50000, 300000),
                    MinStartSellingPrice = "$" + random.Next(40000, 280000),
                    SalesResult = "Pending"
                }),
                TechnologicalComplexList0 = GenerateList(() => new TechnologicalComplex
                {
                    LeasedOut = "Yes",
                    WhenRentedOut = "2022-05-01",
                    Counterparty = "XYZ Corp",
                    RentActivities = "Logistics Center",
                    RentActionPlan = "Long-term lease"
                }),
                AlienationOfAssetsList0 = GenerateList(() => new AlienationOfAssets
                {
                    Completed = "No",
                    AlienationValue = "$" + random.Next(20000, 200000),
                    Other = "N/A",
                    SalesIncome = "$" + random.Next(30000, 250000),
                    FinancialResult = "$" + random.Next(-5000, 50000),
                    MethodOfAlienation = "Auction",
                    ContractDetails = "Contract No. XYZ-123",
                    ConterpartyName = "ABC Ltd.",
                    WriteOffDate = "2024-01-01",
                    Notes = "Pending Approval"
                }),
                CommonComplexProperties = new CommonComplexProperties
                {
                    NameOfAppraiser = "XYZ Appraisers",
                    ComplexReportTaxDate = "2023-01-01",
                    StartTaxCost = "$500,000",
                    ResidualTaxCost = "$250,000",
                    ComplexReportBookkeepingDate = "2023-01-01",
                    StartBookkeepingCost = "$400,000",
                    ResidualBookkeepingCost = "$200,000",
                    ComplexReportMarketDate = "2023-06-01",
                    MarketValueWithDuty = "$600,000",
                    MarketValueWDutyFree = "$580,000",
                    RestrictionsOfRight = "None",
                    MainCharacteristicTypes = "Factory",
                    MainCharacteristicValues = "5000 sq.m",
                    MainCharacteristicUnits = "sq.m",
                    NameNetworks = "Power, Water, Gas",
                    PowerSupply = "Yes",
                    WaterSupply = "Yes",
                    HeatSupply = "Yes",
                    TransportAccessibility = "Highway access",
                    CurrentUsage = "Active Production",
                    PossibleUsage = "Storage, Distribution",
                    TotalBuildingsArea = "5000 sq.m",
                    AreaMeassure = "sq.m",
                    Adress = "123 Industrial Zone, City"
                },
                IncomeAndExpenses = new IncomeAndExpenses
                {
                    PrevYear = "$1,000,000",
                    СurrentYear = "$1,200,000",
                    ActualTotal = "$2,200,000",
                    СurrentYearTotal = "$1,500,000",
                    NotesTotal = "Steady growth observed.",
                    ExpensesList0 = GenerateList(() => new Expenses
                    {
                        Index = "EXP-" + random.Next(1, 100),
                        Name = "Expense " + random.Next(1, 10),
                        FactPrevYear = "$" + random.Next(10000, 50000),
                        FactСurrentYear = "$" + random.Next(20000, 60000),
                        Notes = "Operational Costs"
                    }),
                    IncomeList0 = GenerateList(() => new Income
                    {
                        Index = "INC-" + random.Next(1, 100),
                        Name = "Income " + random.Next(1, 10),
                        FactPrevYear = "$" + random.Next(100000, 500000),
                        FactСurrentYear = "$" + random.Next(200000, 600000),
                        Notes = "Sales Revenue"
                    })
                }
            };
        }

        private static List<T> GenerateList<T>(Func<T> generator)
        {
            var random = new Random();
            int count = random.Next(1, 15);
            List<T> list = new List<T>();
            for (int i = 0; i < count; i++)
            {
                list.Add(generator());
            }
            return list;
        }
    }
