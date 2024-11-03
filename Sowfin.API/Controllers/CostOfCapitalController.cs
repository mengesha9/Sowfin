using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using OfficeOpenXml;
using Sowfin.API.Lib;
using Sowfin.API.ViewModels;
using Sowfin.API.ViewModels.CostOfCapital;
using Sowfin.Data.Abstract;
using Sowfin.Data.Common.Enum;
using Sowfin.Data.Common.Helper;
using Sowfin.Model.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml.Style;
using static Sowfin.API.ViewModels.CostOfCapital.Methods;
using System.Data;
//using System.Runtime.Serialization.Formatters.Binary;
using ExcelDataReader;
using System.Text;
using System.Runtime.Serialization.Formatters.Binary;
using Microsoft.AspNetCore.Http;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using Microsoft.Extensions.Configuration;
using MathNet.Numerics.Statistics;
using System.Diagnostics;

namespace Sowfin.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class CostOfCapitalController : ControllerBase
    {
        private readonly ICostOfCapital iCostOfCapital = null;
        private readonly ICalculateBeta iCalculateBeta = null;
        private readonly IWebHostEnvironment  _hostingEnvironment = null;
        private readonly ISnapshots iSnapshots = null;
        private const string COSTOFCAPITAL = "COST_OF_CAPITAL";
        public IConfiguration Configuration { get; }
        //private radonly ISynonym synonym = null;
        public CostOfCapitalController(IWebHostEnvironment  hostingEnvironment, ICalculateBeta _iCalculateBeta, ICostOfCapital _iCostOfCapital, ISnapshots snapshots, IConfiguration configuration)
        {
            _hostingEnvironment = hostingEnvironment;
            iCalculateBeta = _iCalculateBeta;
            iCostOfCapital = _iCostOfCapital;
            //this.synonym = synonym;
            iSnapshots = snapshots;
            Configuration = configuration;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="UserId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetCostOfCapitals/{UserId}")]
        public ActionResult<Object> GetCostOfCapitals(long UserId)
        {
            ResultObject resultObject = new ResultObject();

            try
            {
                var costOfCapital = iCostOfCapital.FindBy(s => s.UserId == UserId).ToArray();
                CostOfCapital CostOfCapitalObj = new CostOfCapital();
                if (costOfCapital.Length != 0)
                {
                    costOfCapital[0].MarketValueStock =!string.IsNullOrEmpty(costOfCapital[0].MarketValueUnit) ? costOfCapital[0].MarketValueStock / UnitConversion.ReturnDividend(costOfCapital[0].MarketValueUnit) : costOfCapital[0].MarketValueStock;

                    costOfCapital[0].TotalValueStock = !string.IsNullOrEmpty(costOfCapital[0].TotalValueUnit) ? costOfCapital[0].TotalValueStock / UnitConversion.ReturnDividend(costOfCapital[0].TotalValueUnit) : costOfCapital[0].TotalValueStock;

                    costOfCapital[0].MarketValueDebt = !string.IsNullOrEmpty(costOfCapital[0].MarketDebtUnit) ? costOfCapital[0].MarketValueDebt / UnitConversion.ReturnDividend(costOfCapital[0].MarketDebtUnit) : costOfCapital[0].MarketValueDebt;

                    costOfCapital[0].PreferredDividend = !string.IsNullOrEmpty(costOfCapital[0].PreferredDividendUnit) ? costOfCapital[0].PreferredDividend / UnitConversion.ReturnDividend(costOfCapital[0].PreferredDividendUnit) : costOfCapital[0].PreferredDividend;

                    costOfCapital[0].PreferredShare =!string.IsNullOrEmpty(costOfCapital[0].PreferredShareUnit) ? costOfCapital[0].PreferredShare / UnitConversion.ReturnDividend(costOfCapital[0].PreferredShareUnit) : costOfCapital[0].PreferredShare;

                    CostOfCapitalObj = costOfCapital[0];

                    costOfCapital[0].MarketValueStock = Convert.ToDouble(costOfCapital[0].MarketValueStock.ToString("0.##"));
                }

                
               
                CostOfCapitalViewModel costofCapitalVM = new CostOfCapitalViewModel();
                //pause for now
                //var  costofCapitalVM = GenericMapper<CostOfCapital, CostOfCapitalViewModel>.MapObject(CostOfCapitalObj);
                if (CostOfCapitalObj != null)
                {
                    costofCapitalVM.Id = CostOfCapitalObj.Id;
                    costofCapitalVM.MarketValueStock = CostOfCapitalObj.MarketValueStock;
                    costofCapitalVM.TotalValueStock = CostOfCapitalObj.TotalValueStock;
                    costofCapitalVM.MarketValueDebt = CostOfCapitalObj.MarketValueDebt;
                    costofCapitalVM.RiskFreeRate = CostOfCapitalObj.RiskFreeRate;
                    costofCapitalVM.HistoricMarket = CostOfCapitalObj.HistoricMarket;
                    costofCapitalVM.HistoricRiskReturn = CostOfCapitalObj.HistoricRiskReturn;
                    costofCapitalVM.SmallStock = CostOfCapitalObj.SmallStock;
                    costofCapitalVM.RawBeta = CostOfCapitalObj.RawBeta;
                    costofCapitalVM.ApprovalFlag = CostOfCapitalObj.ApprovalFlag;
                    costofCapitalVM.PreferredDividend = CostOfCapitalObj.PreferredDividend;
                    costofCapitalVM.PreferredShare = CostOfCapitalObj.PreferredShare;
                    costofCapitalVM.TaxRate = CostOfCapitalObj.TaxRate;
                    costofCapitalVM.ProjectRisk = CostOfCapitalObj.ProjectRisk;
                    costofCapitalVM.Method = CostOfCapitalObj.Method;
                    costofCapitalVM.MethodType = CostOfCapitalObj.MethodType;
                    costofCapitalVM.UserId = CostOfCapitalObj.UserId;
                    costofCapitalVM.SummaryOutput = CostOfCapitalObj.SummaryOutput;
                    costofCapitalVM.MarketValueUnit = CostOfCapitalObj.MarketValueUnit;
                    costofCapitalVM.TotalValueUnit = CostOfCapitalObj.TotalValueUnit;
                    costofCapitalVM.MarketDebtUnit = CostOfCapitalObj.MarketDebtUnit;
                    costofCapitalVM.PreferredDividendUnit = CostOfCapitalObj.PreferredDividendUnit;
                    costofCapitalVM.PreferredShareUnit = CostOfCapitalObj.PreferredShareUnit;
                    costofCapitalVM.SummaryFlag = CostOfCapitalObj.SummaryFlag;
                    costofCapitalVM.CompanyName = CostOfCapitalObj.CompanyName;
                    costofCapitalVM.CostOfCapitals_Id = CostOfCapitalObj.Id;
                }

                CalculateBeta testCalculateBeta = new CalculateBeta();
                 var testCalculateBetaArray = iCalculateBeta.FindBy(x => x.CostOfCapitals_Id == costofCapitalVM.Id).OrderByDescending(x=>x.Id).ToArray();
                //  testCalculateBeta = iCalculateBeta.GetLatestSingle(x => x.CostOfCapitals_Id == costofCapitalVM.Id).order;
                testCalculateBeta = testCalculateBetaArray.Length !=0 ? testCalculateBetaArray[0] :null;
                if (testCalculateBeta != null)
                {
                    costofCapitalVM.CalculateBeta_Id = testCalculateBeta.Id;
                    costofCapitalVM.CostOfCapitals_Id = testCalculateBeta.CostOfCapitals_Id;
                    costofCapitalVM.Frequency_Id = testCalculateBeta.Frequency_Id;
                    //  costofCapitalVM.Frequency_Value = EnumHelper.DescriptionAttr((FrequencyEnum)testCalculateBeta.Frequency_Id);
                    costofCapitalVM.TargetMarketIndex_Id = testCalculateBeta.TargetMarketIndex_Id;
                    //[
                    costofCapitalVM.TargetRiskFreeRate_Id = testCalculateBeta.TargetRiskFreeRate_Id;
                    costofCapitalVM.TargetRiskFreeRate_Value = EnumHelper.DescriptionAttr((TargetRiskFreeIndexEnum)testCalculateBeta.TargetRiskFreeRate_Id);
                    costofCapitalVM.DataSource_Id = testCalculateBeta.DataSource_Id;
                    costofCapitalVM.DataSource_Value = EnumHelper.DescriptionAttr((DataSourceEnum)testCalculateBeta.DataSource_Id);
                    costofCapitalVM.Duration_FromDate = testCalculateBeta.Duration_FromDate;
                    costofCapitalVM.Duration_toDate = testCalculateBeta.Duration_toDate;
                    costofCapitalVM.Beta_CreatedDate = testCalculateBeta.CreatedDate;
                    costofCapitalVM.Beta_ModifiedDate = testCalculateBeta.ModifiedDate;
                    costofCapitalVM.Beta_Active = testCalculateBeta.Active;
                    costofCapitalVM.BetaValue =Math.Round(Convert.ToDouble(testCalculateBeta.BetaValue),2); 
                }
                else
                    costofCapitalVM.CalculateBeta_Id = 0;

                if (costofCapitalVM == null)
                {
                    resultObject.id = 0;
                    resultObject.result = null;
                    return NotFound(resultObject);
                }
                CostOfCapitalViewModel[] arrayObj = new CostOfCapitalViewModel[1];
                arrayObj[0] = costofCapitalVM;
                return Ok(arrayObj);
            }
            catch (Exception ss)
            {
                Console.WriteLine(ss.Message);
                return BadRequest();
            }
        }

        [HttpGet]
        [Route("test")]
        public ActionResult test()
        {
            try{
                string NotePadPath = Configuration.GetValue<string>("NotePadPath");
                Process.Start(NotePadPath);
                return Ok("DONE");
            }
            catch(Exception ex)
            {
                return BadRequest(ex.Message);
            }

            //Process.Start("D:\\SowFin\\backend\\backend\\Sowfin.API\\wwwroot\\Helper\\notepad.exe");
        }



        [HttpGet]
        [Route("GetBetaValueByCompany")]
        public ActionResult GetBetaValueByCompany(string CompanyName, int FinanceType = 1)
        {

            try
            {
                string Result = string.Empty;
                string ChromeDriverPath = Configuration.GetValue<string>("CromeDriverPath");           
                //string AddressURL = "https://finance.yahoo.com/quote/TCS?p=TCS&.tsrc=fin-srch";           
                if (FinanceType == 1)
                {
                     string AddressURL = string.Empty;
                    AddressURL = "https://finance.yahoo.com/quote/" + CompanyName + "?p=" + CompanyName + "&.tsrc=fin-srch";
                   // AddressURL = "https://finance.yahoo.com/quote/" + CompanyName + "?p=" + CompanyName + "&.tsrc=fin-srch";
                    string path = "//*[@id=Zquote-summaryZ]/div[2]/table/tbody/tr[2]/td[2]/span";
                    path = path.Replace('Z', '"');
                    using (var driver = new ChromeDriver(ChromeDriverPath))// "D:/SowFin/backend/backend/Sowfin.API/Helper
                    {
                        driver.Navigate().GoToUrl(AddressURL);
                        Result = driver.FindElement(By.XPath(path)).Text != "" ? driver.FindElement(By.XPath(path)).Text : "0";
                    }
                }

                return Ok(new  {Result = Result, status = "Success" });

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("GetCostOfCapital/{Id}")]
        public ActionResult<Object> GetCostOfCapital(long Id)
        {
            // if (Id == 0)
            // {
            //     return BadRequest();
            // }
            try
            {
                ResultObject resultObject = new ResultObject();
                var costOfCapital = iCostOfCapital.GetSingle(s => s.Id == Id);
                if (costOfCapital == null )
                {
                    resultObject.id = 0;
                    resultObject.result = 0;
                    return NotFound("No Cost of Capital found");
                }
                return Ok(costOfCapital);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("Evaluate")]
        public ActionResult<Object> AddCostOfCapital([FromBody] CostOfCapitalViewModel model)
        {
            double[] values = {model.MarketValueStock,model.TotalValueStock,model.MarketValueDebt,model.PreferredDividend,
                model.PreferredShare
            };
            string[] units = {model.MarketValueUnit,model.TotalValueUnit,model.MarketDebtUnit,model.PreferredDividendUnit,
                model.PreferredShareUnit
            };

            var output = UnitConversion.ConvertUnits(values, units, 0);
            model.MarketValueStock = output[0];
            model.TotalValueStock = output[1];
            model.MarketValueDebt = output[2];
            model.PreferredDividend = output[3];
            model.PreferredShare = output[4];

            Dictionary<string, double> summaryOutput = NewMethod(model);
            model.SummaryOutput = JsonConvert.SerializeObject(summaryOutput);
            Dictionary<string, object> result = new Dictionary<string, object>
            {
                { "Adjusted Beta (1/3 + 2/3*raw beta)", summaryOutput["AdjustedeBeta"] },
                { "Market Risk Premium", summaryOutput["MarketRiskPremium"] },
                { "Cost of Equity ( rE))", summaryOutput["Costofequity"] },
                { "Cost of Preferred Equity (rP)", summaryOutput["CostofpreferredEquity"] },
                { "Cost of Debt (rD)", summaryOutput["CostOfDebtMethod"] },
                { "Unlevered Cost of Capital/Equity (rU)", summaryOutput["Unleveredcostofcaptial"] },
                { "Weighted Average Cost of Capital (rWACC)", summaryOutput["WeightedAverage"] },
                { "Adjusted WACC", summaryOutput["AdjustedWACC"] }
            };

            try
            {
                if (model.Id == 0)
                {
                    CostOfCapital costOfCapital = new CostOfCapital
                    {
                        // add company to table
                        CompanyName = model.CompanyName,
                        MarketValueStock = model.MarketValueStock,
                        TotalValueStock = model.TotalValueStock,
                        MarketValueDebt = model.MarketValueDebt,

                        RiskFreeRate = model.RiskFreeRate,
                        HistoricMarket = model.HistoricMarket,
                        HistoricRiskReturn = model.HistoricRiskReturn,
                        SmallStock = model.SmallStock,
                        RawBeta = model.RawBeta,

                        ApprovalFlag = model.ApprovalFlag,

                        PreferredDividend = model.PreferredDividend,
                        PreferredShare = model.PreferredShare,

                        TaxRate = model.TaxRate,
                        ProjectRisk = model.ProjectRisk,
                        Method = model.Method,
                        MethodType = model.MethodType,
                        UserId = model.UserId,
                        SummaryOutput = model.SummaryOutput,

                        MarketValueUnit = model.MarketValueUnit,
                        TotalValueUnit = model.TotalValueUnit,
                        MarketDebtUnit = model.MarketDebtUnit,
                        PreferredDividendUnit = model.PreferredDividendUnit,
                        PreferredShareUnit = model.PreferredShareUnit,
                        SummaryFlag = model.SummaryFlag
                    };

                    iCostOfCapital.Add(costOfCapital);
                    iCostOfCapital.Commit();

                    return Ok(result);
                }
                else
                {
                    CostOfCapital updateCostOfCapital = new CostOfCapital
                    {
                        Id = model.Id,
                        CompanyName = model.CompanyName,
                        MarketValueStock = model.MarketValueStock,
                        TotalValueStock = model.TotalValueStock,
                        MarketValueDebt = model.MarketValueDebt,

                        RiskFreeRate = model.RiskFreeRate,
                        HistoricMarket = model.HistoricMarket,
                        HistoricRiskReturn = model.HistoricRiskReturn,
                        SmallStock = model.SmallStock,
                        RawBeta = model.RawBeta,

                        ApprovalFlag = model.ApprovalFlag,

                        PreferredDividend = model.PreferredDividend,
                        PreferredShare = model.PreferredShare,

                        TaxRate = model.TaxRate,
                        ProjectRisk = model.ProjectRisk,
                        Method = model.Method,
                        MethodType = model.MethodType,
                        UserId = model.UserId,
                        SummaryOutput = model.SummaryOutput,

                        MarketValueUnit = model.MarketValueUnit,
                        TotalValueUnit = model.TotalValueUnit,
                        MarketDebtUnit = model.MarketDebtUnit,
                        PreferredDividendUnit = model.PreferredDividendUnit,
                        PreferredShareUnit = model.PreferredShareUnit,
                        SummaryFlag = model.SummaryFlag
                    };

                    iCostOfCapital.Update(updateCostOfCapital);
                    iCostOfCapital.Commit();

                    return Ok(result);
                }
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("DownloadTemplate")]
        public ActionResult<Object> DownloadTemplate([FromBody] CostOfCapitalViewModel model)
        {
            try
            {

                if (model.CalculateBeta_Id != null && model.CalculateBeta_Id == 0)
                    {
                        CalculateBeta CalculateBetaObj = new CalculateBeta
                        {
                            CostOfCapitals_Id = null,
                            MasterId=model.MasterId,
                            FileData = model.FileData,
                            Frequency_Id = model.Frequency_Id,
                            TargetMarketIndex_Id = model.TargetMarketIndex_Id,
                            TargetRiskFreeRate_Id = model.TargetRiskFreeRate_Id,
                            DataSource_Id = model.DataSource_Id,
                            Duration_FromDate = model.Duration_FromDate,
                            Duration_toDate = model.Duration_toDate,
                            ModifiedDate = System.DateTime.UtcNow,
                            CreatedDate = System.DateTime.UtcNow,
                            Active = true
                        };
                        iCalculateBeta.Add(CalculateBetaObj);
                        iCalculateBeta.Commit();
                    }
                    else
                    {
                        CalculateBeta CalculateBetaObj = new CalculateBeta
                        {
                            Id = Convert.ToInt64(model.CalculateBeta_Id),
                            MasterId = model.MasterId,
                            CostOfCapitals_Id = null,
                            FileData = model.FileData,
                            Frequency_Id = model.Frequency_Id,
                            TargetMarketIndex_Id = model.TargetMarketIndex_Id,
                            TargetRiskFreeRate_Id = model.TargetRiskFreeRate_Id,
                            DataSource_Id = model.DataSource_Id,
                            Duration_FromDate = model.Duration_FromDate,
                            Duration_toDate = model.Duration_toDate,
                            ModifiedDate = System.DateTime.Now,
                            Active = model.Beta_Active
                        };
                        iCalculateBeta.Update(CalculateBetaObj);
                        iCalculateBeta.Commit();
                    }

                // generate Excel
                byte[] fileContents;
                var formattedCustomObject = (String)null;
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("CalculateBetaTemplate");

                    // for cell on row1 col1
                    worksheet.Cells[1, 1].Value = "";
                    worksheet.Cells[1, 2].Value = model.CompanyName != null ? model.CompanyName : "";
                    worksheet.Cells[1, 3].Value = EnumHelper.DescriptionAttr((TargetMarketIndexEnum)model.TargetMarketIndex_Id);
                    Console.WriteLine("13  12");
                    worksheet.Cells[1, 4].Value = EnumHelper.DescriptionAttr((TargetRiskFreeIndexEnum)model.TargetRiskFreeRate_Id);
                   Console.WriteLine("13 ");
                    worksheet.Row(1).Style.Border.Top.Style = ExcelBorderStyle.Hair;
                    worksheet.Row(1).Style.Font.Bold = true;
                    worksheet.Row(1).Style.Font.Size = 12;
                    worksheet.Cells[2, 1].Value = "Date";
                    worksheet.Cells[2, 2].Value = "Total Return";
                    worksheet.Cells[2, 3].Value = "Total Return";
                    worksheet.Cells[2, 4].Value = "Total Return";
                    DateTime fromDate = Convert.ToDateTime(model.Duration_FromDate); // Convert.ToDateTime("31/12/2018");
                    DateTime toDate = Convert.ToDateTime(model.Duration_toDate); // Convert.ToDateTime("12/09/2019");  
                    Console.WriteLine("14 ");

                    if ((model.Frequency_Id != null && model.Frequency_Id == (int)FrequencyEnum.Monthly) || (model.Frequency_Value != null && model.Frequency_Value == "Monthly"))
                    {
                        Console.WriteLine("15");
                        List<DateTime> monthslist = GetMonthsBetween(fromDate, toDate);
                        for (int i = 0; i < monthslist.Count; i++)
                        {
                            Console.WriteLine("16");
                            var first = new DateTime(monthslist[i].Date.Year, monthslist[i].Date.Month, 1);
                            var last = first.AddMonths(1).AddDays(-1);
                             Console.WriteLine("17");
                            //remove weekends
                            if(last.DayOfWeek== DayOfWeek.Saturday || last.DayOfWeek == DayOfWeek.Sunday)
                            {
                                last = last.DayOfWeek == DayOfWeek.Saturday ? last.AddDays(-1) : last = last.AddDays(-2); 
                            }

                            worksheet.Cells[i + 3, 1].Value = last.ToShortDateString();
                            // worksheet.Cells[i + 3, 1].Value = monthslist[i].ToShortDateString();
                        }
                    }
                    else if ((model.Frequency_Id != null && model.Frequency_Id == (int)FrequencyEnum.Daily) || (model.Frequency_Value != null && model.Frequency_Value == "Daily"))
                    {
                        List<DateTime> datelist = GetDatesBetween(fromDate, toDate);
                        for (int i = 0; i < datelist.Count; i++)
                        {
                            string test = datelist[i].Date.ToShortDateString();
                            worksheet.Cells[i + 3, 1].Value = datelist[i].ToShortDateString();
                        }
                    }
                    else if ((model.Frequency_Id != null && model.Frequency_Id == (int)FrequencyEnum.Weekly) || (model.Frequency_Value != null && model.Frequency_Value == "Weekly"))
                    {
                        List<DateTime> datelist = GetallFridaysBetween(fromDate, toDate);
                        for (int i = 0; i < datelist.Count; i++)
                        {
                            string test = datelist[i].ToShortDateString();
                            worksheet.Cells[i + 3, 1].Value = datelist[i].ToShortDateString();
                        }
                    }

                    fileContents = package.GetAsByteArray();
                }

                if (fileContents == null || fileContents.Length == 0)
                {
                    return NotFound();
                }
                var inputAsString = Convert.ToBase64String(fileContents);
                formattedCustomObject = JsonConvert.SerializeObject(inputAsString, Formatting.Indented);
                return Ok(formattedCustomObject);
            }
            catch (Exception ss)
            {
                return BadRequest(Convert.ToString("this is the error"+ ss.Message ));
            }
            //return new EmptyResult();
        }

        /// <summary>
        /// get month List Between two dates
        /// </summary>
        /// <param name="from"></param>
        /// <param name="to"></param>
        /// <returns></returns>
        private List<DateTime> GetMonthsBetween(DateTime from, DateTime to)
        {
            to = to.Date.AddMonths(1);
            if (from > to) return GetMonthsBetween(to, from);

            var monthDiff = Math.Abs((to.Year * 12 + (to.Month - 1)) - (from.Year * 12 + (from.Month - 1)));

            if (from.AddMonths(monthDiff) > to || to.Day < from.Day)
            {
                monthDiff -= 1;
            }

            List<DateTime> results = new List<DateTime>();
            for (int i = monthDiff; i >= 1; i--)
            {
                results.Add(to.AddMonths(-i));
            }

            return results;
        }

        private List<DateTime> GetDatesBetween(DateTime startDate, DateTime endDate)
        {
            List<DateTime> allDates = new List<DateTime>();
            for (DateTime date = startDate; date <= endDate; date = date.AddDays(1))
                allDates.Add(date);
            //exclude staurday & sunday
            allDates = allDates.Where(x => x.DayOfWeek != DayOfWeek.Saturday && x.DayOfWeek != DayOfWeek.Sunday).ToList();
            return allDates;
        }

        private List<DateTime> GetallFridaysBetween(DateTime startDate, DateTime endDate)
        {

            List<DateTime> allDates = new List<DateTime>();
            for (DateTime date = startDate; date <= endDate; date = date.AddDays(1))
                allDates.Add(date);
            allDates = allDates.Where(x => x.DayOfWeek == DayOfWeek.Friday).ToList();
            return allDates;

        }

        // Upload Excel
        private DataSet ExcelStreamToDataSet(Stream stream)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            var ds = excelReader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });
            excelReader.Close();
            return ds;
        }        


        [HttpPost]
        [Route("UploadCalculateBeta")]
        public ActionResult UploadCalculateBeta([FromBody]CalculateFileViewModel model)
        {          
            double? beta = null;
            try
            {
                byte[] fileAsBytes = Convert.FromBase64String(model.FileData);
                Stream stream = new MemoryStream(fileAsBytes);
                DataSet DS = this.ExcelStreamToDataSet(stream);
                DataTable dt = DS.Tables[0];

                //Insert Two Columns in dt (Required)
                dt.Columns.Add("Excess Returns", typeof(System.String));
                dt.Columns.Add("Excess2", typeof(System.String));
                DataRow dr1 = dt.Rows[0];
                dr1[4] = "JCP (%)";
                dr1[5] = "S&P 500 (%)";               

                List<double> JCPList = new List<double>();
                List<double> SPList = new List<double>();
                for (int i = 0; i < dt.Rows.Count; i++)
                {

                    if (i > 1)
                    {
                        DataRow drC = dt.Rows[i];
                        DataRow drP = dt.Rows[i - 1];
                        drC[4] = decimal.Round(((Convert.ToDecimal(drC[1]) / Convert.ToDecimal(drP[1])) - 1 - (Convert.ToDecimal(drP[3]) / 1200)) * 100, 2);
                        JCPList.Add(Convert.ToDouble((Convert.ToDouble(drC[1]) / Convert.ToDouble(drP[1])) - 1 - (Convert.ToDouble(drP[3]) / 1200)));
                        drC[5] = decimal.Round(((Convert.ToDecimal(drC[2]) / Convert.ToDecimal(drP[2])) - 1 - (Convert.ToDecimal(drP[3]) / 1200)) * 100, 2);
                        SPList.Add(Convert.ToDouble((Convert.ToDouble(drC[2]) / Convert.ToDouble(drP[2])) - 1 - (Convert.ToDouble(drP[3]) / 1200)));
                    }
                }

                //Average
                //double AvgJcP = Math.Round(JCPList.Average() * 100, 2);
                //double AvgSP = Math.Round(SPList.Average() * 100, 2);

                double AvgJcP = JCPList.Average() * 100;
                double AvgSP = SPList.Average() * 100;

                //manual Standard Deviation Calculation
                double ValJCP = 0;
                double ValSP = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (i > 1)
                    {
                        ValJCP = ValJCP + ((Convert.ToDouble(dt.Rows[i]["Excess Returns"].ToString()) - AvgJcP) * (Convert.ToDouble(dt.Rows[i]["Excess Returns"].ToString()) - AvgJcP));
                        ValSP = ValSP + ((Convert.ToDouble(dt.Rows[i]["Excess2"].ToString()) - AvgSP) * (Convert.ToDouble(dt.Rows[i]["Excess2"].ToString()) - AvgSP));
                    }
                }
                //double SDJCP = Math.Round(Math.Sqrt(ValJCP / ((dt.Rows.Count - 2) - 1)), 2);
                //double SDSP = Math.Round(Math.Sqrt(ValSP / ((dt.Rows.Count - 2) - 1)), 2);
                double SDJCP = Math.Sqrt(ValJCP / ((dt.Rows.Count - 2) - 1));
                double SDSP = Math.Sqrt(ValSP / ((dt.Rows.Count - 2) - 1));


                double correlation = Correlation.Pearson(JCPList.ToArray(), SPList.ToArray());
                beta = correlation * SDJCP / SDSP;

                //Insert Avg,SD,Correlation and beta in datatable
                for (int i = 0; i <= 3; i++)
                {
                    DataRow drCalc = dt.NewRow();
                    if (i == 0)
                    {
                        drCalc[3] = "Average Return (%)";
                        drCalc[4] = Math.Round(AvgJcP, 2);
                        drCalc[5] = Math.Round(AvgSP, 2);
                    }
                    else if (i == 1)
                    {
                        drCalc[3] = "Std. Dev. (%)";
                        drCalc[4] = Math.Round(SDJCP, 2);
                        drCalc[5] = Math.Round(SDSP, 2);
                    }
                    else if (i == 2)
                    {
                        drCalc[3] = "Correlation";
                        drCalc[4] = Math.Round(correlation, 4);
                    }
                    else if (i == 3)
                    {
                        drCalc[3] = "Beta";
                        drCalc[4] = Math.Round(Convert.ToDecimal(beta), 4);
                    }
                    dt.Rows.Add(drCalc);
                }

                //Convert dt to Base64 string 
                byte[] fileContents;
                var formattedCustomObject = (String)null;
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("CalculateBetaTemplate");
                    worksheet.Cells["A1"].LoadFromDataTable(dt, true);

                    int cnt = dt.Rows.Count;
                    worksheet.Cells[1, 1].Value = " ";
                    worksheet.Cells[1, 6].Value = " ";
                    worksheet.Cells[1, 5].Style.Font.Bold = true;
                    worksheet.Cells[2, 5].Style.Font.Bold = true;
                    worksheet.Cells[2, 6].Style.Font.Bold = true;
                    worksheet.Cells[cnt + 1, 4].Style.Font.Bold = true;
                    worksheet.Cells[cnt + 1, 5].Style.Font.Bold = true;
                    worksheet.Cells[cnt, 4].Style.Font.Bold = true;
                    worksheet.Cells[cnt - 1, 4].Style.Font.Bold = true;
                    worksheet.Cells[cnt - 2, 4].Style.Font.Bold = true;

                    fileContents = package.GetAsByteArray();
                    string Base64Str = Convert.ToBase64String(fileContents);
                    formattedCustomObject = JsonConvert.SerializeObject(Base64Str, Formatting.Indented);
                }

                // get dataBy calculateBetaId
                CalculateBeta calculateBetaObj = new CalculateBeta();
                calculateBetaObj = iCalculateBeta.GetSingle(s => s.Id == model.CalculateBeta_Id);

                //save data to beta CalculationTable
                if (calculateBetaObj != null)
                {
                    calculateBetaObj.ModifiedDate = System.DateTime.Now;
                    calculateBetaObj.FileData = formattedCustomObject;
                    calculateBetaObj.BetaValue = Math.Round(Convert.ToDouble(beta), 2);
                    iCalculateBeta.Update(calculateBetaObj);
                    iCalculateBeta.Commit();
                }
                beta = Convert.ToDouble(decimal.Round(Convert.ToDecimal(beta), 2));
            }
            catch (Exception ss)
            {
                Console.WriteLine(ss.Message);
                return BadRequest(ss.Message);

            }
            return Ok(beta);
        }



        [HttpGet]
        [Route("ExportCalculateBeta/{CalculateBeta_Id}")]
        public ActionResult ExportCalculateBeta(long CalculateBeta_Id)
        {
        
            try
            {
                CalculateBeta calculateBetaObj = new CalculateBeta();
                var formattedCustomObject = (string)null;
                calculateBetaObj = iCalculateBeta.GetSingle(s => s.Id == CalculateBeta_Id);
                //save data to beta CalculationTable
                if (calculateBetaObj != null)
                {
                    formattedCustomObject = !string.IsNullOrEmpty(calculateBetaObj.FileData) ? calculateBetaObj.FileData : "No File Exist";
                }
                return Ok(new {result = formattedCustomObject, status = "Success" });

            }
            catch (Exception ss)
            {
                return BadRequest(ss.Message);
            }
        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        private static Dictionary<string, double> NewMethod(CostOfCapitalViewModel model)
        {
            return MathCostOfCapital.SummaryOutput(model);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpDelete]
        [Route("DeleteCostOfCapital/{id}")]
        public ActionResult<Object> DeleteCostOfCapital(long id)
        {
            if (id == 0)
            {
                return BadRequest();
            }

            try
            {
                iCostOfCapital.DeleteWhere(s => s.Id == id);
                return Ok(new { id = id, result = "Successfully deleted" });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("[action]")]
        public ActionResult<Object> ApproveCostOfCapital([FromBody] ApprovalBody model)
        {
            // ApprovalBody model = JsonConvert.DeserializeObject<ApprovalBody>(str);
            try
            {
                if (model.ApprovalFlag == 1)
                {
                    var costOfCapital = iCostOfCapital.FindBy(s => s.UserId == model.UserId).ToArray();
                    if (costOfCapital != null && costOfCapital.Length > 0)
                    {
                        costOfCapital[0].ApprovalFlag = 1;
                        iCostOfCapital.Update(costOfCapital[0]);
                        iCostOfCapital.Commit();
                        return Ok("Successfully approved");
                    }
                    else
                    {
                        return NotFound("record not foud");
                    }
                }
                else
                {
                    return BadRequest("Approval flag not set");
                }
        

            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        // public class ApprovalBody
        // {
        //     public long UserId;
        //     public int ApprovalFlag;
        // }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("[action]")]
        public ActionResult<Object> AddCostOfCapitalSnapshot([FromBody] CostOfCapitalSnapShot model)
        {

            // CostOfCapitalSnapShot model = JsonConvert.DeserializeObject<CostOfCapitalSnapShot>(str);
            try
            {
                Snapshots snapshots = new Snapshots
                {
                    SnapShot = model.SnapShot,
                    Description = model.Description,
                    UserId = model.UserId,
                    SnapShotType = COSTOFCAPITAL
                };

                iSnapshots.Add(snapshots);
                iSnapshots.Commit();
                return Ok(new { id = snapshots.Id, result = "Snapshot saved sucessfully" });
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="UserId"></param>
        /// <returns></returns>

        [HttpGet]
        [Route("CostOfCaptialSnapShots/{UserId}")]
        public ActionResult<Object> CostOfCaptialSnapShots(long UserId)
        {
            try
            {
                var SnapShot = iSnapshots.FindBy(s => s.UserId == UserId && s.SnapShotType == COSTOFCAPITAL);
                if (SnapShot == null)
                {
                    return NotFound("No Snapshots found");
                }
                return Ok(SnapShot);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Id"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("CostOfCapitalSnapShot/{Id}")]
        public ActionResult<Object> CostOfCapitalSnapShot(long Id)
        {
            try
            {
                var SnapShot = iSnapshots.FindBy(s => s.Id == Id && s.SnapShotType == COSTOFCAPITAL);
                if (SnapShot == null)
                {
                    return NotFound("No Snapshots found");
                }
                return Ok(SnapShot);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        // class ResultObject
        // {
        //     public long id;
        //     public object result;
        // }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="UserId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("ExportCostOfCapital/{UserId}")]
        public ActionResult<Object> ExportCostOfCapital(long UserId)
        {
            if (UserId != 0)
            {
                string rootFolder = _hostingEnvironment.WebRootPath;
                string fileName = @"cost_of_capital.xlsx";


                FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
                var formattedCustomObject = (string)null;

                var costOfCapital = iCostOfCapital.FindBy(s => s.UserId == UserId).ToArray();
                if(costOfCapital.Length == 0)
                {
                    return NotFound("No Cost of Capital found");
                }


                using (ExcelPackage package = new ExcelPackage(file))
                {

                    SummaryOutput sumaryOutput = JsonConvert.DeserializeObject<SummaryOutput>(costOfCapital[0].SummaryOutput);
                    var wsCostOfCapital = package.Workbook.Worksheets["CostOfCapital"];

                    wsCostOfCapital = CellValueFormatting("H9", "Market Value Common Stock(E)", "I9", costOfCapital[0].MarketValueStock, costOfCapital[0].MarketValueUnit, 0, 0, wsCostOfCapital);
                    wsCostOfCapital = CellValueFormatting("H10", "Total Value Preferred Stock (P)", "I10", costOfCapital[0].TotalValueStock, costOfCapital[0].TotalValueUnit, 0, 0, wsCostOfCapital);
                    wsCostOfCapital = CellValueFormatting("H11", "Market Value of Net Debt (D)", "I11", costOfCapital[0].MarketValueDebt, costOfCapital[0].MarketDebtUnit, 0, 0, wsCostOfCapital);


                    wsCostOfCapital.Cells["I14"].Value = costOfCapital[0].RiskFreeRate / 100;
                    wsCostOfCapital.Cells["I15"].Value = costOfCapital[0].HistoricMarket / 100;
                    wsCostOfCapital.Cells["I16"].Value = costOfCapital[0].HistoricRiskReturn / 100;
                    wsCostOfCapital.Cells["I17"].Value = costOfCapital[0].SmallStock;
                    wsCostOfCapital.Cells["I18"].Value = costOfCapital[0].RawBeta;
                    wsCostOfCapital.Cells["I19"].Value = sumaryOutput.AdjustedeBeta;
                    wsCostOfCapital.Cells["I19"].Formula = "=1/3+2/3*I18";

                    wsCostOfCapital.Cells["I20"].Value = sumaryOutput.MarketRiskPremium;
                    wsCostOfCapital.Cells["I20"].Formula = "=(I15 - I16) + I17";
                    wsCostOfCapital.Cells["I21"].Value = sumaryOutput.Costofequity;
                    wsCostOfCapital.Cells["I21"].Formula = "= I14 + I19 * I20";
                    wsCostOfCapital.Cells["I24"].Value = costOfCapital[0].PreferredDividend;
                    wsCostOfCapital.Cells["I25"].Value = costOfCapital[0].PreferredShare;
                    wsCostOfCapital = CellValueFormatting("H24", "Preferred Dividend (Dp)", "I24", costOfCapital[0].PreferredDividend, costOfCapital[0].PreferredDividendUnit, 0, 0, wsCostOfCapital);
                    wsCostOfCapital = CellValueFormatting("H25", "Current Preferred Share Price (Pp)", "I25", costOfCapital[0].PreferredShare, costOfCapital[0].PreferredShareUnit, 0, 0, wsCostOfCapital);
                    wsCostOfCapital.Cells["I26"].Value = sumaryOutput.CostofpreferredEquity;
                    wsCostOfCapital.Cells["I26"].Formula = "=I24/I25";
                    if (costOfCapital[0].MethodType == 1)
                    {
                        Method1 results = JsonConvert.DeserializeObject<Method1>(costOfCapital[0].Method);
                        wsCostOfCapital.Cells["I29"].Value = results.RiskFreeRate / 100;
                        wsCostOfCapital.Cells["I30"].Value = results.DefaultSpread / 100;
                        wsCostOfCapital.Cells["I31"].Value = sumaryOutput.CostOfDebtMethod;
                        wsCostOfCapital.Cells["I31"].Formula = "=(I29+I30)";
                        wsCostOfCapital.Cells["I51"].Value = sumaryOutput.CostOfDebtMethod;
                        wsCostOfCapital.Cells["I51"].Formula = "=I29+I30";

                    }
                    else if (costOfCapital[0].MethodType == 2)
                    {
                        Method2 results = JsonConvert.DeserializeObject<Method2>(costOfCapital[0].Method);
                        wsCostOfCapital.Cells["I34"].Value = results.YeildToMaturity / 100;
                        wsCostOfCapital.Cells["I35"].Value = results.ProbabilityOfDefault / 100;
                        wsCostOfCapital.Cells["I36"].Value = results.ExpectedLossRate / 100;
                        wsCostOfCapital.Cells["I37"].Value = sumaryOutput.CostOfDebtMethod;
                        wsCostOfCapital.Cells["I37"].Formula = "=I34-(I35*I36)";
                        wsCostOfCapital.Cells["I51"].Value = sumaryOutput.CostOfDebtMethod;
                        wsCostOfCapital.Cells["I51"].Formula = "=I34-(I35*I36)";

                    }

                    wsCostOfCapital.Cells["I40"].Value = costOfCapital[0].TaxRate / 100;
                    wsCostOfCapital.Cells["I41"].Value = costOfCapital[0].ProjectRisk / 100;

                    wsCostOfCapital.Cells["I47"].Value = sumaryOutput.AdjustedeBeta;
                    wsCostOfCapital.Cells["I47"].Formula = "=1/3+2/3*I18";

                    wsCostOfCapital.Cells["I48"].Value = sumaryOutput.MarketRiskPremium;
                    wsCostOfCapital.Cells["I48"].Formula = "=(I15-I16) + I17";

                    wsCostOfCapital.Cells["I49"].Value = sumaryOutput.Costofequity;
                    wsCostOfCapital.Cells["I49"].Formula = "=I14 + I19 * I20";

                    wsCostOfCapital.Cells["I50"].Value = sumaryOutput.CostofpreferredEquity;
                    wsCostOfCapital.Cells["I50"].Formula = "=I24/I25";


                    if (costOfCapital[0].MethodType == 1)
                    {
                        wsCostOfCapital.Cells["I52"].Value = sumaryOutput.Unleveredcostofcaptial;
                        wsCostOfCapital.Cells["I52"].Formula = "=I21*(I9/(I9+I10+I11)) + I26*(I10/(I9+I10+I11)) + I31*(I11/(I9+I10+I11))";

                        wsCostOfCapital.Cells["I53"].Value = sumaryOutput.WeightedAverage;
                        wsCostOfCapital.Cells["I53"].Formula = "=I21*(I9/(I9+I10+I11)) + I26*(I10/(I9+I10+I11)) + I31*(I11/(I9+I10+I11))*(1-I40)";
                    }
                    else if (costOfCapital[0].MethodType == 2)
                    {
                        wsCostOfCapital.Cells["I52"].Value = sumaryOutput.Unleveredcostofcaptial;
                        wsCostOfCapital.Cells["I52"].Formula = "=I21*(I9/(I9+I10+I11)) + I26*(I10/(I9+I10+I11)) + I37*(I11/(I9+I10+I11))";

                        wsCostOfCapital.Cells["I53"].Value = sumaryOutput.WeightedAverage;
                        wsCostOfCapital.Cells["I53"].Formula = "=I21*(I9/(I9+I10+I11)) + I26*(I10/(I9+I10+I11)) + I37*(I11/(I9+I10+I11))*(1-I40)";
                    }

                    wsCostOfCapital.Cells["I54"].Value = sumaryOutput.AdjustedWACC;
                    wsCostOfCapital.Cells["I54"].Formula = "=I41 + I53";
                    ExcelPackage excelPackage = new ExcelPackage();
                    excelPackage.Workbook.Worksheets.Add("CostOfCapital", wsCostOfCapital);

                    ExcelPackage epOut = excelPackage;
                    //package.Save();
                    byte[] myStream = epOut.GetAsByteArray();
                    var inputAsString = Convert.ToBase64String(myStream);
                    formattedCustomObject = JsonConvert.SerializeObject(inputAsString, Formatting.Indented);
                    excelPackage.Dispose();
                }

                return Ok(formattedCustomObject);
            }
            else
            {
                return BadRequest("Id Not Found ");
            }
        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="cellHeading"></param>
        /// <param name="valueCell"></param>
        /// <param name="value"></param>
        /// <param name="unitForFormat"></param>
        /// <param name="flag"></param>
        /// <param name="formulaRValue"></param>
        /// <param name="wsCapitalStructure"></param>
        /// <returns></returns>
        static ExcelWorksheet CellValueFormatting(string cell, string cellHeading, string valueCell, object value, string unitForFormat,
            int flag, int formulaRValue, ExcelWorksheet wsCapitalStructure)
        {    ///use another for flag check type for using for loop appeding for multiple cells
            if (flag == 0) { wsCapitalStructure.Cells[cell].Value = cellHeading + "($" + unitForFormat.ToUpper() + ")"; }
            else if (flag == 1) { wsCapitalStructure.Cells[cell].Value = cellHeading + "(" + unitForFormat.ToUpper() + ")"; }
            if (formulaRValue == 0) { wsCapitalStructure.Cells[valueCell].Value = value; }
            else if (formulaRValue == 1) { wsCapitalStructure.Cells[valueCell].Formula = Convert.ToString(value); }
            switch (unitForFormat.ToLower())
            {
                case "k":
                    wsCapitalStructure.Cells[valueCell].Style.Numberformat.Format = "$#.00" + "\" K\"";
                    break;
                case "m":
                    wsCapitalStructure.Cells[valueCell].Style.Numberformat.Format = "$#,,.00" + "\" M\"";
                    break;
                case "b":
                    wsCapitalStructure.Cells[valueCell].Style.Numberformat.Format = "$#,,,.00" + "\" B\"";
                    break;
                case "t":
                    wsCapitalStructure.Cells[valueCell].Style.Numberformat.Format = "$#,,,,.00" + "\" T\"";
                    break;
            }
            return wsCapitalStructure;
        }

        //[HttpGet]
        //[Route("SaveIntoPG/{Id}")]
        //public async Task<IActionResult> SaveIntoPG(int Id)
        //{

        //    string rootFolder = _hostingEnvironment.WebRootPath;
        //    string fileName = @"INCOME.xlsx";
        //    FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
        //    string finfield = null, category = null, othertag = null, synonymss = "";
        //    List<string> array = new List<string>();
        //    using (ExcelPackage package = new ExcelPackage(file))
        //    {
        //        var ws = package.Workbook.Worksheets["Sheet1"];
        //        finfield = ws.Cells["A" + Id].Value.ToString();
        //        category = ws.Cells["B" + Id].Value.ToString();
        //        othertag = ws.Cells["C" + Id].Value.ToString();
        //        var syn = ws.Cells["E" + Id].Value;
        //        if (syn != null)
        //        {
        //            synonymss = syn.ToString();
        //        }
        //        if (synonymss.Contains(","))
        //        {
        //            array = synonymss.Split(",").ToList<string>();
        //            List<SynonymTable> synonymTables = new List<SynonymTable>();
        //            if (array.Count != 0)
        //            {
        //                for (int i = 0; i < array.Count; i++)
        //                {
        //                    SynonymTable synonyms = new SynonymTable
        //                    {
        //                        FinField = finfield,
        //                        Category = category,
        //                        OtherTags = othertag,
        //                        Synonym = array[i].Trim(),
        //                        StatementType = "BALANCE_SHEET"

        //                    };
        //                    synonymTables.Add(synonyms);
        //                }

        //            }

        //            synonym.SaveMany(synonymTables);
        //        }
        //        else
        //        {
        //            SynonymTable synonyms = new SynonymTable
        //            {
        //                FinField = finfield,
        //                Category = category,
        //                OtherTags = othertag,
        //                Synonym = synonymss.Trim(),
        //                StatementType = "BALANCE_SHEET"

        //            };
        //            synonym.Save(synonyms);

        //        }
        //    }
        //    return Ok(new
        //    {
        //        finfield,
        //        category,
        //        othertag,
        //        array
        //    });
    }
}

