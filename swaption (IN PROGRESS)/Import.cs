using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Globalization;
using System.Diagnostics;
using System.Linq;
using System.Text;
using TotalRisk.ExcelWrapper;
using TotalRisk.Utilities;
using TotalRisk.MortgageModel;

using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel.Application;
using XlFileFormat = Microsoft.Office.Interop.Excel.XlFileFormat;

namespace TotalRisk.ValuationModule
{
    public enum CashFlowType
    {
        RiskNeutral = (int)0,
        RiskRente = (int)+1
    }
    public class Import_ValuationModule : CImport
    {

        // General data:
        public ScenarioList ReadScenarios(string fileName)
        {
            const string SheetName = "Scenario's";
            object[,] values = null;
            if (File.Exists(fileName))
            {
                try
                {
                    values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, SheetName, "A1");
                }
                catch (Exception exc)
                {
                    throw new ApplicationException("Fout tijdens inlezen van scenario's uit sheet " + SheetName + " in bestand " + fileName);
                }
            }
            ScenarioList scenarios = ScenarioList.getInstance(values);
            return scenarios;
        }
        public ScopeData ReadFileScopeData(string fileName)
        {
            ScopeData data = new ScopeData();
            if (File.Exists(fileName))
            {
                object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, "Blad1", "A1");
                data = ScopeData.getInstance(values);
            }
            return data;
        }
        public SortedList<string, double> ReadFund_Participations_CurrencyHedged(object[,] values)
        {
            // FundScoupe | OTSO Scope | Participaton
            SortedList<string, double> data = null;
            Dictionary<string, int> headers = HeaderNamesColumns(values); // import headers
            data = new SortedList<string, double>();
            for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
            {
                string sExternHedged = ReadFieldAsString(values, row, 0, headers, "Intern/Extern hedged").Trim().ToUpper();
                if (sExternHedged != "EXTERN")
                {
                    continue;
                }
                string sFund_ID = ScopeData.getScopeFormated(ReadFieldAsString(values, row, 0, headers, "Security ID"));
                double fHedgePercentage = ReadFieldAsDouble(values, row, 0, headers, "Percentage");
                if (!data.ContainsKey(sFund_ID))
                {
                    data.Add(sFund_ID, fHedgePercentage);
                }
            }
            return data;
        }

        // Spaarlos data:
        public Dictionary<string, CurveList> Read_Spaarlos_Curves(string fileName, DateTime dtNow, ErrorList errors)
        {
            const string SheetName = "rentecurves per bron_jaar";
            ScenarioList scenarios = new ScenarioList();
            // Load Positions:
            const int RowHeader = 1;
            const int RowStart = 2;
            bool invalidDateReported = false;
            Dictionary<string, CurveList> SpaarlosCurves = new Dictionary<string, CurveList>();
            CurveList curveList = new CurveList();
            if (File.Exists(fileName))
            {
                try
                {
                    object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, SheetName, "A1");
                    Dictionary<string, int> colNames = HeaderNamesColumns(values);
                    int N = values.GetUpperBound(DimensionRow);
                    for (int row = RowStart; row <= N; row++)
                    {
                        try
                        {
                            if (values[row, 1] == null)
                            {
                                continue;
                            }
                            DateTime? dtReport = ReadFieldAsDateTime(values, row, 0, colNames, "ReportDate");
                            if (dtReport != dtNow)
                            {
                                string message = "RapportageDatum van kasstroom ongelijk aan rapportage datum in ScenarioTool voor positie in regel " + row.ToString();
                                if (!invalidDateReported)
                                {
                                    if (MessageBox.Show(message + ". Bestand alsnog verwerken?", "Ongeldige rapportagedatum", MessageBoxButtons.OKCancel) == DialogResult.OK)
                                    {
                                        errors.AddWarning("Ongeldige rapportagedatums in bestand. Gebruiker heeft melding genegeerd");
                                    }
                                    else
                                    {
                                        errors.AddError("Ongeldige rapportagedatums in bestand. Gebruiker heeft verwerking gestopt");
                                        return SpaarlosCurves;
                                    }

                                    invalidDateReported = true;
                                }
                                errors.AddWarning(message);
                            }
                            string curveName = ReadFieldAsString(values, row, 0, colNames, "Curve name");
                            string curveCode = ReadFieldAsString(values, row, 0, colNames, "ASR Unique Curve Code").ToUpper().Trim();
                            Curve curve = new Curve();
                            int col = HeaderNameId(colNames, "Rating") + 1;
                            int M = values.GetUpperBound(DimensionCol);
                            while ((col <= M) && (values[RowHeader, col] != null))
                            {
                                string maturityString = ReadFieldAsString(values, RowHeader, col).ToUpper();
                                int tenor = int.Parse(maturityString.Replace("Y", ""));
                                double rate = ReadFieldAsDouble(values, row, col);
                                CurvePoint point = new CurvePoint(tenor, rate);
                                curve.AddPoint(point);
                                col++;
                            }
                            curveList.Add(curveCode, curve);
                        }
                        catch (Exception exc)
                        {
                            errors.AddError("Fout tijdens inlezen vastrentend positie in regel " + row.ToString() + "\n" + exc.Message);
                        }
                    }

                }
                catch (Exception exc)
                {
                    throw new ApplicationException("Fout tijdens inlezen van Spaarlos Curves uit sheet " + SheetName + " in bestand " + fileName);
                }
                finally
                {
                }

            }
            SpaarlosCurves.Add("EUR", curveList);
            return SpaarlosCurves;
        }
        public PositionList Read_Spaarlos_Positions(DateTime dtNow, string fileName, ScenarioList scenarios,
            Dictionary<string, CurveList> zeroDiscountCurves, bool bRiskMargin, ErrorList errors)
        {
            // Defne the base curves:
            CurveList scenarioZeroCurves = new CurveList();
            Curve zeroCurve, scenarioZeroCurve;
            string ccy;
            TotalRisk.Utilities.Scenario baseScenario = scenarios.getScenarioFairValue();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
            {
                ccy = scenarioCurve.m_sName;
                scenarioZeroCurve = scenarioCurve.m_Curve;
                scenarioZeroCurves.Add(ccy.ToUpper(), scenarioZeroCurve);
            }

            // Load Positions:
            const int RowHeader = 1;
            const int RowStart = 2;

            bool invalidDateReported = false;
            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadUsedRangeValues(fileName, "kasstromen per maand", "A1");
            Dictionary<string, int> colNames = HeaderNamesColumns(values);
            PositionList positions = new PositionList();
            for (int row = RowStart; row <= values.GetUpperBound(DimensionRow); row++)
            {
                try
                {
                    if (values[row, 1] == null)
                    {
                        continue;
                    }
                    string dataSource = ReadFieldAsString(values, row, 0, colNames, "Purpose").Trim().ToUpper();
                    if (dataSource.ToLower() != "spaarlos")
                    {
                        continue;
                    }

                    DateTime? dtReport = ReadFieldAsDateTime(values, row, 0, colNames, "ReportDate");
                    if (dtReport != dtNow)
                    {
                        string message = "ReportDate van kasstroom ongelijk aan rapportage datum in ScenarioTool voor positie in regel " + row.ToString();
                        if (!invalidDateReported)
                        {
                            if (MessageBox.Show(message + ". Bestand alsnog verwerken?", "Ongeldige rapportagedatum", MessageBoxButtons.OKCancel) == DialogResult.OK)
                            {
                                errors.AddWarning("Ongeldige rapportagedatums in bestand. Gebruiker heeft melding genegeerd");
                            }
                            else
                            {
                                errors.AddError("Ongeldige rapportagedatums in bestand. Gebruiker heeft verwerking gestopt");
                                return positions;
                            }

                            invalidDateReported = true;
                        }
                        errors.AddWarning(message);
                    }
                    string sActuariesScenarioName = "Best estimate";
                    string sActuariesScenarioID = "";
                    if (HeaderNameExists(colNames, "Actuaries Scenarion Name"))
                    {
                        sActuariesScenarioName = ReadFieldAsString(values, row, 0, colNames, "Actuaries Scenarion Name");
                    }
                    if (HeaderNameExists(colNames, "Actuaries Scenarion ID"))
                    {
                        sActuariesScenarioID = ReadFieldAsString(values, row, 0, colNames, "Actuaries Scenarion ID");
                    }
                    if (!bRiskMargin)
                    {
                        if ("best estimate" != sActuariesScenarioName.ToLower())
                        {
                            continue;
                        }
                    }
                    CashflowSchedule sched = new CashflowSchedule();
                    int col = HeaderNameId(colNames, "Boekwaarde") + 1;
                    while ((col <= values.GetUpperBound(DimensionCol)) && (values[RowHeader, col] != null))
                    {
                        string bucketString = ReadFieldAsString(values, RowHeader, col).ToUpper().Trim();
                        bool isMonthly = bucketString.EndsWith("M");
                        DateTime? dtBucket = dtReport.Value;
                        if (isMonthly)
                        {
                            int period = int.Parse(bucketString.Replace("M", ""));
                            dtBucket = dtReport.Value.AddMonths(period);
                        }
                        else
                        {
                            int period = int.Parse(bucketString.Replace("Y", ""));
                            dtBucket = dtReport.Value.AddYears(period);
                        }
                        if (values[row, col] != null) // "Cash Flow" line
                        {
                            Cashflow cf = new Cashflow(dtBucket.Value, (double)values[row, col]);
                            sched.Add(cf);
                        }
                        col++;
                    }
                    Position position = new Position();
                    position.m_sActuariesScenarioName = sActuariesScenarioName;
                    position.m_sActuariesScenarioID = sActuariesScenarioID;
                    if (HeaderNameExists(colNames, "Scope"))
                    {
                        position.m_sScope3 = ReadFieldAsString(values, row, 0, colNames, "Scope");
                    }
                    if (HeaderNameExists(colNames, "ISSUER SCOPE"))
                    {
                        position.m_sScope3_Issuer = ReadFieldAsString(values, row, 0, colNames, "ISSUER SCOPE");
                    }
                    if (HeaderNameExists(colNames, "INVERTOR SCOPE"))
                    {
                        position.m_sScope3_Investor = ReadFieldAsString(values, row, 0, colNames, "INVERTOR SCOPE");
                    }
                    if (HeaderNameExists(colNames, "CIC Code"))
                    {
                        position.m_sCIC_LL = ReadFieldAsString(values, row, 0, colNames, "CIC Code");
                        position.m_sCIC_SCR = position.m_sCIC_LL;
                        position.m_sCIC = position.m_sCIC_LL;
                    }
                    position.m_sBalanceType = "Assets";
                    position.m_sGroup = "Fixed Income";
                    if (HeaderNameExists(colNames, "modelpunt"))
                    {
                        position.m_sPortfolioId = ReadFieldAsString(values, row, 0, colNames, "modelpunt");
                    }
                    position.m_sAccount = ReadFieldAsString(values, row, 0, colNames, "RDS-STA account");
                    position.m_sAccount_LL = position.m_sAccount;
                    position.m_sSMS_Entity = ReadFieldAsString(values, row, 0, colNames, "SMS entiteit");

                    string collateralType = ReadFieldAsString(values, row, 0, colNames, "Onderpand");
                    string securityType = ReadFieldAsString(values, row, 0, colNames, "Spaar/Hybride");
                    string securityType_LL = ReadFieldAsString(values, row, 0, colNames, "Security Type");
                    string securityId = ReadFieldAsString(values, row, 0, colNames, "Uniek nummer");
                    string securityName = ReadFieldAsString(values, row, 0, colNames, "Security name");
                    string securityCurrency = ReadFieldAsString(values, row, 0, colNames, "Currency");
                    double volume = ReadFieldAsDouble(values, row, 0, colNames, "Boekwaarde");
                    double fxRate = ReadFieldAsDouble(values, row, 0, colNames, "FX rate");
                    double collateral = 0;
                    if (HeaderNameExists(colNames, "Collateral coverage percent"))
                    {
                        collateral = ReadFieldAsDouble(values, row, 0, colNames, "Collateral coverage percent");
                    }
                    position.m_sRow = row.ToString();
                    position.m_sSecurityType = securityType;
                    position.m_sSecurityType_LL = securityType_LL;
                    position.m_sSecurityID_LL = securityId;
                    position.m_sSecurityName_LL = securityName;
                    position.m_sPositionId = position.m_sSecurityID_LL;
                    position.m_sLegId = "0";
                    position.m_sCurrency = securityCurrency;
                    position.m_fFxRate = fxRate;
                    position.m_fVolume = volume;
                    position.m_fCollateralCoveragePercentage = collateral;
                    position.m_sCollateralType = collateralType;
                    position.m_sIssuerCreditQuality = ReadFieldAsString(values, row, 0, colNames, "Tegenpartij SII Credit Quality Step");
                    position.m_sCounterpartyIssuer_Name = ReadFieldAsString(values, row, 0, colNames, "Tegenpartij");
                    position.m_sCounterpartyGroup_Name = ReadFieldAsString(values, row, 0, colNames, "Hoofdpartij");
                    position.m_sCounterparty_Name = position.m_sCounterpartyIssuer_Name;
                    position.m_sCounterparty_LEI = ReadFieldAsString(values, row, 0, colNames, "Hoofdtegenpartij LEI Code");
                    position.m_sSelectieIndex_LL = "1000000001001000"; // rate, valuta, counterpaty
                    if (HeaderNameExists(colNames, "Termijncontract"))
                    {
                        bool bCounterpartyRisk = ReadFieldAsBool(values, row, 0, colNames, "Termijncontract");
                        if (!bCounterpartyRisk && 0 == position.m_fCollateralCoveragePercentage)
                        {
                            position.m_sSelectieIndex_LL = "1000000001110000"; // rate, valuta, spread, concentration
                        }
                    }


                    position.m_bEEA = true;
                    position.m_sCountryCurrency = position.m_sCurrency;
                    position.m_bGovGuarantee = false;
                    if (position.m_sIssuerCreditQuality == "" || position.m_sIssuerCreditQuality.Substring(0, 1) == "NR")
                    {//  seven as NR
                        position.m_dIssuerCreditQuality = 7;
                    }
                    else
                    {
                        position.m_dIssuerCreditQuality = Convert.ToInt32(position.m_sIssuerCreditQuality.Substring(0, 1));
                    }
                    string discountCurveName = ReadFieldAsString(values, row, 0, colNames, "ASR Unique Curve Code").Trim().ToUpper();
                    string curr = position.m_sCurrency;
                    string curveCode = discountCurveName;
                    if (!zeroDiscountCurves.ContainsKey(position.m_sCurrency))
                    {
                        curr = "EUR";
                    }
                    if (!zeroDiscountCurves[position.m_sCurrency].ContainsKey(discountCurveName))
                    {
                        curveCode = "SWAP";
                    }
                    Curve discountCurve = zeroDiscountCurves[curr][curveCode];

                    Instrument_Cashflow instrument = new Instrument_Cashflow(dtNow, position.m_sSecurityType_LL, sched);
                    instrument.m_sCouponType = "FIXED";

                    string ccyTranslated = Position.TranslateCurrency_Curve(scenarioZeroCurves, position.m_sCurrency);
                    zeroCurve = scenarioZeroCurves[ccyTranslated];
                    instrument.Init(dtNow, zeroCurve, discountCurve, curveCode); // Spread will be floored at -10%

                    position.m_Instrument = instrument;
                    position.m_sDATA_Source = dataSource;
                    positions.AddPosition(position);

                    if (dtReport > instrument.m_MaturityDate)
                    {
                        position.m_bHasMessage = true;
                        position.m_sMessage = " Warning : Instrument is matured. Fair Value wordt 0 voor alle scenarios.";
                        errors.Add("Row = " + position.m_sRow + " ID = " + position.m_sPositionId + position.m_sMessage);
                    }
                    else
                    {
                        if ((instrument.m_fDirtyValue != 0) && (instrument.m_CashFlowSchedule.Count() == 0))
                        {
                            position.m_bHasMessage = true;
                            position.m_sMessage = " Warning : CleanValue is niet nul, maar er zijn geen kasstromen. Fair Value wordt 0 voor alle scenarios.";
                            errors.Add("Row = " + position.m_sRow + " ID = " + position.m_sPositionId + position.m_sMessage);
                        }
                        else if (double.IsNaN(instrument.m_fImpliedSpread))
                        {
                            position.m_bHasMessage = true;
                            position.m_sMessage = " Warning : Spread kan niet worden bepaald. Fair Value wordt gebruikt voor alle scenarios.";
                            errors.Add("Row = " + position.m_sRow + " ID = " + position.m_sPositionId + position.m_sMessage);
                        }
                    }
                    System.Windows.Forms.Application.DoEvents();
                }
                catch (Exception exc)
                {
                    errors.AddError("Fout tijdens inlezen vastrentend positie in regel " + row.ToString() + "\n" + exc.Message);
                }
            }

            return positions;
        }
        // cash Positions:
        public PositionList ReadCashPositions_GARC(DateTime dtNow, string fileName, ErrorList errors)
        {
            PositionList positions = new PositionList();

            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, "DATA", "A1");
            // Read column names into dictionary to map column name to column number
            Dictionary<string, int> columnNames = HeaderNamesColumns(values);

            for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
            {
                string sCIC_ID_LL = ReadFieldAsString(values, row, 0, columnNames, "Cic Id (Laagste Lt Level)");
                bool bLookThroughData = false;
                Instrument_Cash_OriginalData entry = new Instrument_Cash_OriginalData();
                entry.m_sBalanceType = "Assets";
                entry.m_sGroup = "Cash";

                entry.m_dtReport = ReadFieldAsDateTime(values, row, 0, columnNames, "Reporting Date");
                entry.m_dtStartDate = ReadFieldAsDateTime(values, row, 0, columnNames, "Start date");
                entry.m_dtEndDate = ReadFieldAsDateTime(values, row, 0, columnNames, "End date");
                entry.m_bIsOvernight = ReadFieldAsBool(values, row, 0, columnNames, "Overnight");

                entry.m_sSelectieIndex_LL = "0000000001001000"; // valuta (10), counterparty (13)
                entry.m_sDataSource = ReadFieldAsString(values, row, 0, columnNames, "DATA Source");
                entry.m_bLookThroughData = bLookThroughData;
                entry.m_sAccount = ReadFieldAsString(values, row, 0, columnNames, "Tagetik Account");
                entry.m_sAccount_LL = ReadFieldAsString(values, row, 0, columnNames, "SAP account");
                entry.m_sPortfolioID = ReadFieldAsString(values, row, 0, columnNames, "Portfolio Id");

                entry.m_sScope3 = ReadFieldAsString(values, row, 0, columnNames, "Scope 3 code");
                entry.m_sScope3 = ScopeData.getScopeFormated(entry.m_sScope3);
                entry.m_sScope3_Issuer = ReadFieldAsString(values, row, 0, columnNames, "Scope code issuer");
                entry.m_sScope3_Issuer = ScopeData.getScopeFormated(entry.m_sScope3_Issuer);
                entry.m_sScope3_Investor = ReadFieldAsString(values, row, 0, columnNames, "Scope code investor ");
                entry.m_sScope3_Investor = ScopeData.getScopeFormated(entry.m_sScope3_Investor);

                entry.m_sCIC_ID = sCIC_ID_LL;
                entry.m_sCIC_ID_LL = sCIC_ID_LL;
                entry.m_sSecurity_Name_LL = ReadFieldAsString(values, row, 0, columnNames, "Security Name (Laagste Lt Level)");
                entry.m_sSecurity_ID_LL = ReadFieldAsString(values, row, 0, columnNames, "Security Id (Laagste Lt Level)");
                entry.m_sSecurity_Name = entry.m_sSecurity_Name_LL;
                entry.m_sSecurity_ID = entry.m_sSecurity_ID_LL;
                entry.m_sSecurity_Type_LL = ReadFieldAsString(values, row, 0, columnNames, "Instrument Type");

                entry.m_sCurrency = ReadFieldAsString(values, row, 0, columnNames, "Currency (Laagste Lt Level)").ToUpper().Trim();
                entry.m_sCountryCode = entry.m_sCurrency;
                entry.m_sCurrencyCountry = entry.m_sCurrency;

                entry.m_fMarketValue_EUR = ReadFieldAsDouble(values, row, 0, columnNames, "Marktwaarde (Euro)");
                entry.m_fNominal = ReadFieldAsDouble(values, row, 0, columnNames, "Nominale waarde (Euro)");
                entry.m_fCoupon = ReadFieldAsDouble(values, row, 0, columnNames, "Coupon") / 100;
                entry.m_fCollateralPerc = ReadFieldAsDouble(values, row, 0, columnNames, "Collateral coverage");
                entry.m_fImpairmentValue_PC = 0;
                entry.m_fImpairedCostValue_PC = 0;
                entry.m_dType = ReadFieldAsInt(values, row, 0, columnNames, "Type 1/2");
                entry.m_fFxRate = 1.0 / ReadFieldAsDouble(values, row, 0, columnNames, "FX Rate  LL");

                entry.m_sGroupCounterpartyName = ReadFieldAsString(values, row, 0, columnNames, "Counterparty Groep: lei name").ToUpper().Trim();
                entry.m_sGroupCounterpartyLEI = ReadFieldAsString(values, row, 0, columnNames, "Counterparty Groep: lei code").ToUpper().Trim();
                entry.m_sGroupCounterpartyCQS = ReadFieldAsString(values, row, 0, columnNames, "Counterparty Groep: Credit Quality Step").ToUpper().Trim();
                if (entry.m_sGroupCounterpartyCQS == "" || entry.m_sGroupCounterpartyCQS.Substring(0, 1) == "NR")
                {//  seven as NR
                    entry.m_dGroupCounterpartyCQS = 7;
                }
                else
                {
                    entry.m_dGroupCounterpartyCQS = Convert.ToInt32(entry.m_sGroupCounterpartyCQS.Substring(0, 1));
                }

                Instrument_Cash instrument = new Instrument_Cash(entry);
                Position position = new Position();
                position.m_sDATA_Source = entry.m_sDataSource;
                position.m_sRow = row.ToString();
                position.m_bIsLookThroughPosition = entry.m_bLookThroughData;

                position.m_Instrument = instrument;
                position.m_sSelectieIndex_LL = entry.m_sSelectieIndex_LL;
                position.m_sBalanceType = entry.m_sBalanceType;
                position.m_sGroup = entry.m_sGroup;
                position.m_sScope3 = entry.m_sScope3;
                position.m_sScope3_Issuer = entry.m_sScope3_Issuer;
                position.m_sScope3_Investor = entry.m_sScope3_Investor;
                position.m_bICO = false;
                position.m_sUniquePositionId = entry.m_sSecurity_ID_LL;
                position.m_fVolume = entry.m_fNominal;
                position.m_fFairValue = entry.m_fMarketValue_EUR;
                position.m_sSecurityType_LL = "CASH";
                position.m_sCIC = entry.m_sCIC_ID;
                position.m_sCIC_LL = entry.m_sCIC_ID_LL;
                position.m_sCIC_SCR = entry.m_sCIC_ID_LL;
                position.m_sPortfolioId = entry.m_sPortfolioID;
                position.m_sPositionId = entry.m_sSecurity_ID_LL;
                position.m_sSecurityID_LL = entry.m_sSecurity_ID_LL;
                position.m_sSecurityName_LL = entry.m_sSecurity_Name_LL;
                position.m_sCurrency = entry.m_sCurrency;
                position.m_sCountryCurrency = entry.m_sCurrency;
                position.m_sDATA_Source = entry.m_sDataSource;
                position.m_sAccount = entry.m_sAccount;
                position.m_sAccount_LL = entry.m_sAccount_LL;
                position.m_sECAP_Category_LL = entry.m_sECAP_Category_LL;
                position.m_fCollateralCoveragePercentage = entry.m_fCollateralPerc;
                positions.AddPosition(position);
            }

            return positions;
        }
        public PositionList ReadCashPositions_IMW(DateTime dtNow, string fileName, ErrorList errors)
        {
            PositionList positions = new PositionList();

            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, "Integrale aanlevering IMW maand", "A1");
            // Read column names into dictionary to map column name to column number
            Dictionary<string, int> columnNames = HeaderNamesColumns(values);

            for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
            {
                string sBalanceType = "Assets";
                string reportCode = ReadFieldAsString(values, row, 0, columnNames, "SelectieIndex LL");
                if (reportCode.Substring(9, 1) != "1") // 10-th digit for valuta report
                {
                    continue;
                }
                bool nonCash = true;
                string sCIC_ID_LL = ReadFieldAsString(values, row, 0, columnNames, "Cic Id Ll");
                string CIC3_LL = sCIC_ID_LL.Substring(2, 1);
                string CIC34_LL = sCIC_ID_LL.Substring(2, 2);
                string sSecurityType = ReadFieldAsString(values, row, 0, columnNames, "Security Type Ll").Trim().ToUpper();
                if ("79" == CIC34_LL && "CALL MONEY" == sSecurityType)
                {
                    nonCash = false;
                    reportCode = reportCode.Substring(0, 12) + "0" + reportCode.Substring(13);
                    double mktValue = ReadFieldAsDouble(values, row, 0, columnNames, "Market Value EUR LL");
                    if (mktValue < 0)
                    {
                        sBalanceType = "Liabilities";
                    }
                } else if ("71" == CIC34_LL || "72" == CIC34_LL)
                {
                    nonCash = false;
                } else if ("24" == CIC34_LL && "STR REPO" == sSecurityType)
                {
                    nonCash = false;
                }

                if (nonCash)
                {
                    continue;
                }
                string sPortfolioPurpose = ReadFieldAsString(values, row, 0, columnNames, "Portfolio-purpose").Trim().ToUpper();
                if ("FUNDING" == sPortfolioPurpose)
                {
                    continue;
                }
                if ("PARTLOAN" == sSecurityType)
                {
                    continue;
                }

                string sCIC_ID = ReadFieldAsString(values, row, 0, columnNames, "CIC Id");
                string CIC3 = sCIC_ID.Substring(2, 1);
                bool bLookThroughData = false;
                if ("4" == CIC3)
                {
                    bLookThroughData = true;
                }
                Instrument_Cash_OriginalData entry = new Instrument_Cash_OriginalData();
                entry.m_sBalanceType = sBalanceType;
                entry.m_sGroup = "Cash";

                entry.m_dtReport = ReadFieldAsDateTime2(values, row, 0, columnNames, "Reporting Date");

                entry.m_sSelectieIndex_LL = reportCode;
                entry.m_sDataSource = "IMW";
                entry.m_bLookThroughData = bLookThroughData;
                entry.m_sAccount = ReadFieldAsString(values, row, 0, columnNames, "RDS-STA account");
                entry.m_sAccount_LL = ReadFieldAsString(values, row, 0, columnNames, "RDS-STA account LT");

                entry.m_sScope3 = ReadFieldAsString(values, row, 0, columnNames, "Ecs Cons Ecap asr");
                entry.m_sScope3 = ScopeData.getScopeFormated(entry.m_sScope3);

                entry.m_sCIC_ID = sCIC_ID;
                entry.m_sCIC_ID_LL = sCIC_ID_LL;
                entry.m_sSecurity_ID = ReadFieldAsString(values, row, 0, columnNames, "Security Id");
                entry.m_sSecurity_ID_LL = ReadFieldAsString(values, row, 0, columnNames, "Security Id Ll");
                entry.m_sSecurity_Name = ReadFieldAsString(values, row, 0, columnNames, "Security Name");
                entry.m_sSecurity_Name_LL = ReadFieldAsString(values, row, 0, columnNames, "Security Name Ll");
                entry.m_sSecurity_Type = ReadFieldAsString(values, row, 0, columnNames, "Security Type");
                entry.m_sSecurity_Type_LL = ReadFieldAsString(values, row, 0, columnNames, "Security Type Ll");
                entry.m_sPortfolioID = ReadFieldAsString(values, row, 0, columnNames, "Portfolio Id");
                entry.m_sCountryCode = ReadFieldAsString(values, row, 0, columnNames, "Country Code Laagste Lt Level");
                entry.m_sCurrencyCountry = ReadFieldAsString(values, row, 0, columnNames, "Country Currency Ll").ToUpper().Trim();
                entry.m_sCurrency = ReadFieldAsString(values, row, 0, columnNames, "Currency Laagste Lt Level").ToUpper().Trim();
                double FXRate;
                string test = ReadFieldAsString(values, row, 0, columnNames, "Fx Rate Qc Pc Laagste Lt Level");
                if (test == "")
                {
                    FXRate = 1;
                }
                else
                {
                    FXRate = 1.0 / ReadFieldAsDouble(values, row, 0, columnNames, "Fx Rate Qc Pc Laagste Lt Level");
                }
                if (FXRate <= 0)
                {
                    FXRate = 1;
                }
                entry.m_fFxRate = 1.0 / FXRate; // Foreign Curr in EUR
                entry.m_fNominal = ReadFieldAsDouble(values, row, 0, columnNames, "Balnomval qc");
                entry.m_fCollateralPerc = ReadFieldAsDouble(values, row, 0, columnNames, "Coll Coverage Laagste Level");
                entry.m_fMarketValue_EUR = ReadFieldAsDouble(values, row, 0, columnNames, "Market Value EUR LL");


                entry.m_bICO = ReadFieldAsBool(values, row, 0, columnNames, "Eliminatie ASR");
                entry.m_bIsOvernight = true;
                entry.m_dType = 1;
                entry.m_fCoupon = ReadFieldAsDouble(values, row, 0, columnNames, "Coupon Perc Laagste Lt Level") / 100;

                entry.m_fImpairmentValue_PC = 0;
                entry.m_fImpairedCostValue_PC = 0;
                // credit quality steps Group Vounterparty:
                entry.m_sGroupCounterpartyName = ReadFieldAsString(values, row, 0, columnNames, "Groep tegenpartij naam Ll");
                entry.m_sGroupCounterpartyLEI = ReadFieldAsString(values, row, 0, columnNames, "Groep tegenpartij LEI Ll");
                test = ReadFieldAsString(values, row, 0, columnNames, "Groep tegenpartij Credit Quality Step Ll");
                entry.m_sGroupCounterpartyCQS = test;
                if (test == "" || test.Substring(0, 1) == "NR")
                {//  seven as NR
                    entry.m_dGroupCounterpartyCQS = 7;
                }
                else
                {
                    entry.m_dGroupCounterpartyCQS = Convert.ToInt32(test.Substring(0, 1));
                }

                //
                test = ReadFieldAsString(values, row, 0, columnNames, "Initial Start date").Trim();
                if (test == "")
                {
                    entry.m_dtStartDate = entry.m_dtReport.Value;
                }
                else
                {
                    entry.m_dtStartDate = ReadFieldAsDateTime2(values, row, 0, columnNames, "Initial Start date");
                }
                test = ReadFieldAsString(values, row, 0, columnNames, "Maturity LL").Trim();
                if (test == "")
                {
                    entry.m_dtEndDate = entry.m_dtReport.Value.AddDays(1);
                }
                else
                {
                    entry.m_dtEndDate = ReadFieldAsDateTime2(values, row, 0, columnNames, "Maturity LL");
                }


                Instrument_Cash instrument = new Instrument_Cash(entry);
                Position position = new Position();
                position.m_sDATA_Source = entry.m_sDataSource;
                position.m_sRow = row.ToString();
                position.m_bIsLookThroughPosition = entry.m_bLookThroughData;

                position.m_Instrument = instrument;
                position.m_sSelectieIndex_LL = entry.m_sSelectieIndex_LL;
                position.m_sBalanceType = entry.m_sBalanceType;
                position.m_sGroup = entry.m_sGroup;
                position.m_sScope3 = entry.m_sScope3;
                position.m_sScope3_Issuer = entry.m_sScope3_Issuer;
                position.m_sScope3_Investor = entry.m_sScope3_Investor;
                position.m_bICO = entry.m_bICO;
                position.m_sUniquePositionId = entry.m_sSecurity_ID_LL;
                position.m_fVolume = entry.m_fNominal;
                position.m_fFairValue = entry.m_fMarketValue_EUR;
                position.m_sSecurityType_LL = entry.m_sSecurity_Type_LL;
                position.m_sCIC = entry.m_sCIC_ID;
                position.m_sCIC_LL = entry.m_sCIC_ID_LL;
                position.m_sCIC_SCR = entry.m_sCIC_ID_LL;
                position.m_sPortfolioId = entry.m_sPortfolioID;
                position.m_sPositionId = entry.m_sSecurity_ID_LL;
                position.m_sSecurityID = entry.m_sSecurity_ID;
                position.m_sSecurityID_LL = entry.m_sSecurity_ID_LL;
                position.m_sSecurityName_LL = entry.m_sSecurity_Name_LL;
                position.m_sCurrency = entry.m_sCurrency;
                position.m_sCountryCurrency = entry.m_sCurrency;
                position.m_sDATA_Source = entry.m_sDataSource;
                position.m_sAccount = entry.m_sAccount;
                position.m_sAccount_LL = entry.m_sAccount_LL;
                position.m_sECAP_Category_LL = entry.m_sECAP_Category_LL;
                position.m_fCollateralCoveragePercentage = entry.m_fCollateralPerc;
                positions.AddPosition(position);
            }

            return positions;
        }

        // Equity data:
        public PositionList ReadEquityPositions_IMW(DateTime dtNow, string IMW_FileName, ErrorList errors)
        {
            TimeSpan ts = dtNow - new DateTime(2015, 12, 31);
            double age = ts.TotalDays / 365.0;
            double OMA_Correction_Fund = Math.Pow(0.9, age);
            PositionList positions = new PositionList();
            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(IMW_FileName, "Integrale aanlevering IMW maand", "A1");
            // Read column names into dictionary to map column name to column number
            Dictionary<string, int> columnNames = HeaderNamesColumns(values);

            for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
            {
                string reportCode = ReadFieldAsString(values, row, 0, columnNames, "SelectieIndex LL");
                if (reportCode.Substring(4, 1) != "1") // 5-th digit for Equity report
                {
                    continue;
                }
                double OMA_Correction = 1;
                string sCIC_ID = ReadFieldAsString(values, row, 0, columnNames, "CIC Id");
                string CIC3 = sCIC_ID.Substring(2, 1);
                string CIC34 = sCIC_ID.Substring(2, 2);
                bool bLookThroughData = false;
                if ("4" == CIC3)
                {
                    bLookThroughData = true;
                    OMA_Correction = OMA_Correction_Fund;
                }
                string sCIC_ID_LL = ReadFieldAsString(values, row, 0, columnNames, "Cic Id Ll");
                string CIC3_LL = sCIC_ID_LL.Substring(2, 1);
                string CIC34_LL = sCIC_ID_LL.Substring(2, 2);
                bool nonEquity = true;
                if (bLookThroughData)
                {
                    if ("4" == CIC3_LL ||
                        "5" == CIC3_LL ||
                        "6" == CIC3_LL ||
                        "A" == CIC3_LL ||
                        "B" == CIC3_LL ||
                        "C" == CIC3_LL ||
                        "D" == CIC3_LL ||
                        "E" == CIC3_LL ||
                        "F" == CIC3_LL
                        )
                    {
                        nonEquity = false;
                    }
                    else if ("22" == CIC34_LL ||
                             "31" == CIC34_LL ||
                             "32" == CIC34_LL ||
                             "33" == CIC34_LL ||
                             "39" == CIC34_LL ||
                             "09" == CIC34_LL
                        )
                    {
                        nonEquity = false;
                    }
                }
                else
                {
                    if ("5" == CIC3_LL ||
                        "6" == CIC3_LL
                        )
                    {
                        nonEquity = false;
                    }
                    else if ("22" == CIC34_LL ||
                             "31" == CIC34_LL ||
                             "32" == CIC34_LL ||
                             "33" == CIC34_LL ||
                             "39" == CIC34_LL ||
                             "09" == CIC34_LL
                        )
                    {
                        nonEquity = false;
                    }
                }
                if (nonEquity)
                {
                    //                    continue;
                }
                Instrument_Equity_OriginalData entry = new Instrument_Equity_OriginalData();
                entry.m_sBalanceType = "Assets";
                entry.m_sGroup = "Shares";

                //                double dateValue = ReadFieldAsDouble(values, row, 0, colNames, "Reporting Date");
                //                contractData.m_dtReport = getExcelDateFromDoubleDate(dateValue);
                entry.m_dtReport = ReadFieldAsDateTime2(values, row, 0, columnNames, "Reporting Date");

                entry.m_sSelectieIndex_LL = reportCode;
                entry.m_sDataSource = "IMW";
                entry.m_bLookThroughData = bLookThroughData;
                entry.m_sAccount = ReadFieldAsString(values, row, 0, columnNames, "RDS-STA account");
                entry.m_sAccount_LL = ReadFieldAsString(values, row, 0, columnNames, "RDS-STA account LT");

                entry.m_sScope3 = ReadFieldAsString(values, row, 0, columnNames, "Ecs Cons Ecap asr");
                entry.m_sScope3 = ScopeData.getScopeFormated(entry.m_sScope3);

                entry.m_sCIC_ID = sCIC_ID;
                entry.m_sCIC_ID_LL = sCIC_ID_LL;
                entry.m_sSecurity_ID = ReadFieldAsString(values, row, 0, columnNames, "Security Id");
                entry.m_sSecurity_ID_LL = ReadFieldAsString(values, row, 0, columnNames, "Security Id Ll");
                entry.m_sSecurity_Name = ReadFieldAsString(values, row, 0, columnNames, "Security Name");
                entry.m_sSecurity_Name_LL = ReadFieldAsString(values, row, 0, columnNames, "Security Name Ll");
                entry.m_sSecurity_Type = ReadFieldAsString(values, row, 0, columnNames, "Security Type");
                entry.m_sSecurity_Type_LL = ReadFieldAsString(values, row, 0, columnNames, "Security Type Ll");

                entry.m_sBenchmark_ECAP = ReadFieldAsString(values, row, 0, columnNames, "Benchmark ECAP Ll");
                entry.m_sECAP_Category_LL = ReadFieldAsString(values, row, 0, columnNames, "ECAP Category Ll");
                entry.m_sCountryCode = ReadFieldAsString(values, row, 0, columnNames, "Country Code Laagste Lt Level");
                entry.m_sCurrencyCountry = ReadFieldAsString(values, row, 0, columnNames, "Country Currency Ll").ToUpper().Trim();
                entry.m_sCurrency = ReadFieldAsString(values, row, 0, columnNames, "Currency Laagste Lt Level").ToUpper().Trim();
                entry.m_fNominal = ReadFieldAsDouble(values, row, 0, columnNames, "Balnomval qc");
                if (HeaderNameExists(columnNames, "Balcostval Impaired PC LL"))
                {
                    entry.m_fBalcostval_Impaired_PC_LL = ReadFieldAsDouble(values, row, 0, columnNames, "Balcostval Impaired PC LL");
                }
                entry.m_fImpairmentValue_PC = ReadFieldAsDouble(values, row, 0, columnNames, "Impairment Value PC");
                entry.m_fImpairedCostValue_PC = ReadFieldAsDouble(values, row, 0, columnNames, "Balcostval Impaired PC");
                entry.m_sParticipation_Fm = ReadFieldAsString(values, row, 0, columnNames, "Participation Fm").ToUpper().Trim();

                if ("22" == CIC34_LL)
                {// Convertible Bonds
                    entry.m_MarketValue_EUR = ReadFieldAsDouble(values, row, 0, columnNames, "Convertible Optie Waarde Ll");
                }
                else
                {
                    entry.m_MarketValue_EUR = ReadFieldAsDouble(values, row, 0, columnNames, "Market Value Eur Ll");
                }
                entry.m_fPercentageOMA = ReadFieldAsDouble(values, row, 0, columnNames, "Oma Perc");
                entry.m_dType = 2;
                string temp = ReadFieldAsString(values, row, 0, columnNames, "Type 1 or 2 Ll").Trim();
                if ("1" == temp)
                {
                    entry.m_dType = 1;
                }
                entry.m_sPortfolioID = ReadFieldAsString(values, row, 0, columnNames, "Portfolio Id");
                entry.m_fFxRate = 1.0 / ReadFieldAsDouble(values, row, 0, columnNames, "Fx Rate Qc Pc Laagste Lt Level"); // Foreign Curr in EUR

                entry.m_bICO = ReadFieldAsBool(values, row, 0, columnNames, "Eliminatie ASR");
                entry.m_bICO = false; // 2019M08: temporally, later we should read the Vastgoed file to get them back for ASR NL.
                entry.m_sInfrastructureInvestmentType = ReadFieldAsString(values, row, 0, columnNames, "Infra type LT").ToUpper().Trim();

                Instrument_Equity instrument = new Instrument_Equity(entry, OMA_Correction);
                Position position = new Position();
                position.m_sDATA_Source = entry.m_sDataSource;
                position.m_sRow = row.ToString();
                position.m_bIsLookThroughPosition = entry.m_bLookThroughData;

                position.m_Instrument = instrument;
                position.m_sSelectieIndex_LL = entry.m_sSelectieIndex_LL;
                position.m_sBalanceType = entry.m_sBalanceType;
                position.m_sGroup = entry.m_sGroup;
                position.m_sScope3 = entry.m_sScope3;
                position.m_bICO = entry.m_bICO;
                position.m_fVolume = entry.m_fNominal;
                position.m_fFairValue = entry.m_MarketValue_EUR;
                position.m_sSecurityType_LL = ("" != entry.m_sSecurity_Type_LL) ? entry.m_sSecurity_Type_LL : entry.m_sSecurity_Type;
                position.m_sCIC = entry.m_sCIC_ID;
                position.m_sCIC_LL = entry.m_sCIC_ID_LL;
                position.m_sCIC_SCR = "NL31";
                position.m_sPortfolioId = entry.m_sPortfolioID;
                position.m_sPositionId = entry.m_sSecurity_ID_LL;
                position.m_sSecurityID_LL = entry.m_sSecurity_ID_LL;
                position.m_sSecurityID = entry.m_sSecurity_ID;
                position.m_sSecurityName_LL = entry.m_sSecurity_Name_LL;
                position.m_sCurrency = entry.m_sCurrency;
                position.m_sCountryCurrency = entry.m_sCurrency;
                position.m_fFxRate = entry.m_fFxRate;
                position.m_sDATA_Source = entry.m_sDataSource;
                position.m_sAccount = entry.m_sAccount;
                position.m_sAccount_LL = entry.m_sAccount_LL;
                position.m_sECAP_Category_LL = entry.m_sECAP_Category_LL;
                positions.AddPosition(position);
            }

            return positions;
        }
        public PositionList ReadEquityPositions_Participations(DateTime dtNow, string Participations_FileName, ErrorList errors)
        {
            TimeSpan ts = dtNow - new DateTime(2015, 12, 31);
            double age = ts.TotalDays / 365.0;
            double OMA_Correction_Fund = Math.Pow(0.9, age);
            OMA_Correction_Fund = 1; // for participations have fixed OMA perentage
            PositionList positions = new PositionList();

            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(Participations_FileName, "DATA", "A1");
            // Read column names into dictionary to map column name to column number
            Dictionary<string, int> columnNames = HeaderNamesColumns(values);

            for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
            {
                string sCIC_ID = ReadFieldAsString(values, row, 0, columnNames, "CIC Id");
                string CIC3 = sCIC_ID.Substring(2, 1);
                string CIC34 = sCIC_ID.Substring(2, 2);
                bool bLookThroughData = false;
                string sCIC_ID_LL = sCIC_ID;
                string CIC3_LL = sCIC_ID_LL.Substring(2, 1);
                string CIC34_LL = sCIC_ID_LL.Substring(2, 2);
                Instrument_Equity_OriginalData entry = new Instrument_Equity_OriginalData();
                entry.m_sBalanceType = "Assets";
                entry.m_sGroup = "Shares";

                //                double dateValue = ReadFieldAsDouble(values, row, 0, colNames, "Reporting Date");
                //                contractData.m_dtReport = getExcelDateFromDoubleDate(dateValue);
                entry.m_dtReport = ReadFieldAsDateTime(values, row, 0, columnNames, "Reporting Date");

                entry.m_sSelectieIndex_LL = "0000100001010000"; // equity (5), valuta(10), concentration(12)
                entry.m_sDataSource = "Participations";
                entry.m_bLookThroughData = bLookThroughData;
                entry.m_sAccount = ReadFieldAsString(values, row, 0, columnNames, "Tagetik Account");
                entry.m_sAccount_LL = entry.m_sAccount;

                entry.m_sScope3 = ReadFieldAsString(values, row, 0, columnNames, "Scope");
                entry.m_sScope3 = ScopeData.getScopeFormated(entry.m_sScope3);
                entry.m_sScope3_Issuer = ReadFieldAsString(values, row, 0, columnNames, "Scope Issuer");
                entry.m_sScope3_Issuer = ScopeData.getScopeFormated(entry.m_sScope3_Issuer);
                entry.m_sScope3_Investor = ReadFieldAsString(values, row, 0, columnNames, "Scope Investor");
                entry.m_sScope3_Investor = ScopeData.getScopeFormated(entry.m_sScope3_Investor);

                entry.m_sCIC_ID = sCIC_ID;
                entry.m_sCIC_ID_LL = sCIC_ID_LL;
                entry.m_sSecurity_Name = ReadFieldAsString(values, row, 0, columnNames, "Data source");
                entry.m_sSecurity_Name_LL = ReadFieldAsString(values, row, 0, columnNames, "Security Name");
                entry.m_sSecurity_ID = ReadFieldAsString(values, row, 0, columnNames, "Security Id");
                entry.m_sSecurity_ID_LL = entry.m_sSecurity_ID;
                entry.m_sSecurity_Type_LL = ReadFieldAsString(values, row, 0, columnNames, "Equity Type SCR");

                entry.m_sBenchmark_ECAP = ReadFieldAsString(values, row, 0, columnNames, "Equity Type ECAP");
                entry.m_sECAP_Category_LL = entry.m_sBenchmark_ECAP;
                entry.m_sCurrency = ReadFieldAsString(values, row, 0, columnNames, "Currency").ToUpper().Trim();
                entry.m_sCountryCode = entry.m_sCurrency;
                entry.m_sCurrencyCountry = entry.m_sCurrency;

                entry.m_MarketValue_EUR = ReadFieldAsDouble(values, row, 0, columnNames, "Market Value");
                entry.m_fNominal = entry.m_MarketValue_EUR;
                entry.m_fImpairmentValue_PC = 0;
                entry.m_fImpairedCostValue_PC = 0;
                entry.m_sParticipation_Fm = "";
                //                entry.m_fNominal = ReadFieldAsDouble(values, row, 0, columnNames, "Balnomval qc");
                //                entry.m_fImpairmentValue_PC = ReadFieldAsDouble(values, row, 0, columnNames, "Impairment Value PC");
                //                entry.m_fImpairedCostValue_PC = ReadFieldAsDouble(values, row, 0, columnNames, "Impaired Cost Value PC");
                //                entry.m_sParticipation_Fm = ReadFieldAsString(values, row, 0, columnNames, "Participation Fm").ToUpper().Trim();

                entry.m_fPercentageOMA = ReadFieldAsDouble(values, row, 0, columnNames, "OMA");
                switch (entry.m_sSecurity_Type_LL.ToLower())
                {
                    case "strategisch":
                        entry.m_dType = 3;
                        break;
                    case "type1":
                        entry.m_dType = 1;
                        break;
                    case "type2":
                        entry.m_dType = 2;
                        break;
                    case "quinf":
                        entry.m_dType = 4;
                        break;
                    case "quinfc":
                        entry.m_dType = 5;
                        break;
                    default:
                        entry.m_dType = 2;
                        break;
                }
                entry.m_sPortfolioID = entry.m_sScope3;
                entry.m_fFxRate = 1.0 / ReadFieldAsDouble(values, row, 0, columnNames, "Fx Rate");
                entry.m_bICO = false; // to be implemented

                Instrument_Equity instrument = new Instrument_Equity(entry, OMA_Correction_Fund);
                Position position = new Position();
                position.m_sDATA_Source = entry.m_sDataSource;
                position.m_sRow = row.ToString();
                position.m_bIsLookThroughPosition = entry.m_bLookThroughData;

                position.m_Instrument = instrument;
                position.m_sSelectieIndex_LL = entry.m_sSelectieIndex_LL;
                position.m_sBalanceType = entry.m_sBalanceType;
                position.m_sGroup = entry.m_sGroup;
                position.m_sScope3 = entry.m_sScope3;
                position.m_sScope3_Issuer = entry.m_sScope3_Issuer;
                position.m_sScope3_Investor = entry.m_sScope3_Investor;
                position.m_bICO = entry.m_bICO;
                position.m_fVolume = entry.m_fNominal;
                position.m_fFairValue = entry.m_MarketValue_EUR;
                position.m_sSecurityType_LL = "SHARE";
                position.m_sCIC = entry.m_sCIC_ID;
                position.m_sCIC_LL = entry.m_sCIC_ID_LL;
                position.m_sCIC_SCR = "NL31";
                position.m_sPortfolioId = entry.m_sPortfolioID;
                position.m_sPositionId = entry.m_sSecurity_ID_LL;
                position.m_sSecurityID_LL = entry.m_sSecurity_ID_LL;
                position.m_sSecurityName_LL = entry.m_sSecurity_Name_LL;
                position.m_sCurrency = entry.m_sCurrency;
                position.m_sCountryCurrency = entry.m_sCurrency;
                position.m_sDATA_Source = entry.m_sDataSource;
                position.m_sAccount = entry.m_sAccount;
                position.m_sAccount_LL = entry.m_sAccount_LL;
                position.m_sECAP_Category_LL = entry.m_sECAP_Category_LL;
                positions.AddPosition(position);
            }

            return positions;
        }
        public PositionList ReadSharePositions_FullBal(DateTime dtNow, string fileNameBase, ErrorList errors,
            out SortedList<string, SortedList<string, double>> percentageList_ECAP, out SortedList<string, SortedList<string, double>> percentageList_SCR)
        {
            const int ColScope = 1;
            const int ColValue = 2;

            PositionList positions = new PositionList();
            percentageList_ECAP = new SortedList<string, SortedList<string, double>>();
            percentageList_SCR = new SortedList<string, SortedList<string, double>>();
            List<string> IndexList = new List<string>();

            int rowTotal, RowStart, RowEnd, RowIndexName;
            try
            {

                // read the ECAP percentages per scope in:
                object[,] percentages = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileNameBase, "AAND+HF+PE per BM", "A2"); // first table
                RowIndexName = 2;
                rowTotal = 1;
                while (percentages[rowTotal, ColScope] != null)
                {
                    rowTotal++;
                }
                RowStart = 3;
                RowEnd = rowTotal - 1;
                for (int row = RowStart; row <= RowEnd; row++)
                {
                    SortedList<string, double> list = new SortedList<string, double>();
                    for (int col = ColValue; col <= percentages.GetUpperBound(DimensionCol); col++)
                    {
                        if (percentages[RowIndexName, col] == null)
                            continue;
                        string index = ReadFieldAsString(percentages, RowIndexName, col);
                        if (!IndexList.Contains(index))
                        {
                            IndexList.Add(index);
                        }
                        double percentage = ReadFieldAsDouble(percentages, row, col);
                        list.Add(index, percentage);
                    }
                    string percentageScope = ReadFieldAsString(percentages, row, ColScope);
                    percentageList_ECAP.Add(percentageScope, list);
                }
                // read the SCR percentages per scope in:
                percentages = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileNameBase, "AAND+HF+PE per BM", "A21"); // second table
                RowIndexName = 2;
                rowTotal = 1;
                while (percentages[rowTotal, ColScope] != null)
                {
                    rowTotal++;
                }
                RowStart = 3;
                RowEnd = rowTotal - 1;
                for (int row = RowStart; row <= RowEnd; row++)
                {
                    SortedList<string, double> list = new SortedList<string, double>();
                    for (int col = ColValue; col <= percentages.GetUpperBound(DimensionCol); col++)
                    {
                        if (percentages[RowIndexName, col] == null)
                            continue;

                        string index = ReadFieldAsString(percentages, RowIndexName, col);
                        if (!IndexList.Contains(index))
                        {
                            throw new ApplicationException("New SCR Index = " + index + " which is not present in ECAP indexes!");
                        }
                        double percentage = ReadFieldAsDouble(percentages, row, col);
                        list.Add(index, percentage);
                    }
                    string percentageScope = ReadFieldAsString(percentages, row, ColScope);
                    percentageList_SCR.Add(percentageScope, list);
                }
                int numberOfIndexes = IndexList.Count;
                // read share amounts per scope:
                object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileNameBase, "AAND+HF+PE", "A2");
                RowIndexName = 2;
                rowTotal = 1;
                while (values[rowTotal, ColScope] != null)
                {
                    rowTotal++;
                }
                RowStart = 3;
                RowEnd = rowTotal - 1;
                for (int row = RowStart; row <= RowEnd; row++)
                {
                    foreach (string indexName in IndexList)
                    {
                        Position position = new Position();
                        position.m_sBalanceType = "Assets";
                        position.m_sGroup = "Shares";
                        position.m_sScope3 = ReadFieldAsString(values, row, ColScope);
                        position.m_fVolume = ReadFieldAsDouble(values, row, ColValue);
                        position.m_sSecurityType_LL = "SHARE_" + indexName.ToUpper().Trim();
                        position.m_sCIC_LL = "NL31";
                        position.m_sCIC_SCR = "NL31";
                        position.m_sPortfolioId = position.m_sScope3;
                        position.m_sPositionId = position.m_sScope3 + "_" + position.m_sSecurityType_LL;
                        position.m_sSecurityID_LL = indexName;
                        position.m_sSecurityName_LL = indexName;
                        position.m_sDATA_Source = "AANDELEN_FULLBAL";
                        positions.AddPosition(position);
                    }
                }
            }
            catch (Exception exc)
            {
                errors.AddError("Fout tijdens inlezen aadelen positie " + "\n" + exc.Message);
            }

            return positions;
        }
        public PositionList ReadOptionPositions_IMW(DateTime dtNow, string fileName, ScenarioList scenarios, ErrorList errors, OptionMarketDataList marketDataList)
        {
            const int RowStart = 2;
            // Defne the base curves:
            CurveList scenarioZeroCurves = new CurveList();
            Curve zeroCurve, scenarioZeroCurve;
            string ccy;
            TotalRisk.Utilities.Scenario baseScenario = scenarios.getScenarioFairValue();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
            {
                ccy = scenarioCurve.m_sName;
                scenarioZeroCurve = scenarioCurve.m_Curve;
                scenarioZeroCurves.Add(ccy.ToUpper(), scenarioZeroCurve);
            }
            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, "Integrale aanlevering IMW maand", "A1");
            Dictionary<string, int> colNames = HeaderNamesColumns(values);
            PositionList positions = new PositionList(values.GetUpperBound(DimensionRow));
            int year, month, day, row = 0;
            bool isEquityOption = false;
            try
            {
                for (row = RowStart; row <= values.GetUpperBound(DimensionRow); row++)
                {
                    string reportCode = ReadFieldAsString(values, row, 0, colNames, "SelectieIndex LL");
                    if (reportCode.Substring(7, 1) != "1")
                    {
                        continue;
                    }

                    string cic_ID = ReadFieldAsString(values, row, 0, colNames, "Cic Id Ll");
                    string cic_ID_last2 = cic_ID.Substring(2, 2);
                    if (cic_ID_last2 == "B1" || cic_ID_last2 == "C1")
                    {
                        isEquityOption = true;
                    }
                    else
                    {
                        isEquityOption = false;
                    }
                    if (!isEquityOption) continue;
                    Instrument_EquityOption_OriginalData entry = new Instrument_EquityOption_OriginalData();
                    entry.m_sEquityIndexName["SCR"] = "Equity_Average"; // type 1, but scenario Average
                    entry.m_sEquityIndexName["ECAP"] = "Equity_VM_EURO"; // 
                    entry.m_Type = cic_ID_last2.StartsWith("C") ? EquityOptionType.Put : EquityOptionType.Call;
                    entry.m_fStrike = ReadFieldAsDouble(values, row, 0, colNames, "Strike Price Laagste Level");
                    double doubleDateFormat = ReadFieldAsDouble(values, row, 0, colNames, "Maturity Call LL");
                    year = (int)(doubleDateFormat / 10000);
                    month = (int)((doubleDateFormat - year * 10000) / 100);
                    day = (int)(doubleDateFormat - year * 10000 - month * 100);
                    entry.m_MaturityDate = DateTimeExtensions.New(year, month, day).Value;
                    entry.m_sSecurityID_LL = ReadFieldAsString(values, row, 0, colNames, "Security Id Ll");
                    entry.m_sSecurityName_LL = ReadFieldAsString(values, row, 0, colNames, "Security Name Ll");

                    entry.m_fDividendYield = ReadFieldAsDouble(values, row, 0, colNames, "Underlying DividendYield");
                    entry.m_fUnderlyingPrice = ReadFieldAsDouble(values, row, 0, colNames, "Underlying MarketValue");
                    OptionMarketData marketData = marketDataList[entry.m_sSecurityName_LL];
                    entry.m_fDividendYield = marketData.DividendYield;
                    entry.m_fUnderlyingPrice = marketData.UnderlyingPrice;

                    Position position = new Position();
                    position.m_sDATA_Source = "IMW";
                    position.m_sRow = row.ToString();
                    position.m_bIsLookThroughPosition = false;

                    position.m_sBalanceType = "Assets";
                    position.m_sGroup = "Put Options";

                    string portfolioID = ReadFieldAsString(values, row, 0, colNames, "Portfolio Id");
                    position.m_fFairValue = ReadFieldAsDouble(values, row, 0, colNames, "Market Value Eur Ll");
                    position.m_fFxRate = 1.0 / ReadFieldAsDouble(values, row, 0, colNames, "Fx Rate Qc Pc Laagste Lt Level");
                    position.m_fVolume = ReadFieldAsDouble(values, row, 0, colNames, "BalNomVal LT");
                    //                    EquityOptionType type = ReadFieldAsString(values, row, 0, colNames, "Derivaten").ToUpper().StartsWith("C") ? EquityOptionType.Call : EquityOptionType.Put;
                    position.m_sSelectieIndex_LL = reportCode;
                    position.m_sCIC_LL = cic_ID;
                    position.m_sCIC_SCR = position.m_sCIC_LL;
                    position.m_sSecurityType_LL = ReadFieldAsString(values, row, 0, colNames, "Instrument Type Ll").ToUpper();
                    position.m_sScope3 = ReadFieldAsString(values, row, 0, colNames, "Ecs Cons Ecap asr");
                    position.m_sScope3 = ScopeData.getScopeFormated(position.m_sScope3);
                    position.m_sPortfolioId = portfolioID;
                    string securityType = ReadFieldAsString(values, row, 0, colNames, "Security Type Ll").ToUpper();
                    position.m_sSecurityID_LL = entry.m_sSecurityID_LL;
                    position.m_sSecurityName_LL = entry.m_sSecurityName_LL;
                    position.m_sPositionId = position.m_sSecurityID_LL;

                    //                    string UnderlyingEquityName = ReadFieldAsString(values, row, 0, colNames, "Underlying Index Ll").ToLower();
                    //                    string ScenarioName = "EQ index 1"; // until 2020M11
                    //                    position.m_sNameScenarioEquityIndex = ScenarioName.ToLower(); // until 2020M11
                    position.m_sNameScenarioEquityIndex = "Equity_Average"; // type 1, but scenario Average

                    Instrument_EquityOption instrument = new Instrument_EquityOption(entry);
                    position.m_fSCR_weight = Math.Max(0, Math.Min(1, instrument.Expiry(dtNow)));

                    string ccyTranslated = Position.TranslateCurrency_Curve(scenarioZeroCurves, position.m_sCurrency);
                    zeroCurve = scenarioZeroCurves[ccyTranslated];

                    instrument.Init(dtNow, zeroCurve, position.m_fFairValue / position.m_fVolume);

                    position.m_Instrument = instrument;

                    position.m_sSecurityID = ReadFieldAsString(values, row, 0, colNames, "Security Id");
                    position.m_sCIC = ReadFieldAsString(values, row, 0, colNames, "Cic Id").ToUpper();
                    position.m_fAccruedInterest_LL = ReadFieldAsDouble(values, row, 0, colNames, "Accrued Interest Ll");
                    position.m_fAccruedDividend_LL = ReadFieldAsDouble(values, row, 0, colNames, "Accrued Dividend Ll");
                    position.m_sAccount = ReadFieldAsString(values, row, 0, colNames, "RDS-STA account");
                    position.m_sAccount_LL = ReadFieldAsString(values, row, 0, colNames, "RDS-STA account LT");
                    position.m_sECAP_Category_LL = ReadFieldAsString(values, row, 0, colNames, "ECAP Category Ll");

                    positions.AddPosition(position);

                    if (double.IsNaN(instrument.m_fImpliedVol))
                    {
                        errors.Add("Warning " + position.m_sPositionId + " : Implied volatility kan niet worden bepaald. Fair Value wordt gebruikt voor alle scenarios.");
                    }

                    System.Windows.Forms.Application.DoEvents();
                }
            }
            catch (Exception exc)
            {
                errors.AddError("Fout tijdens inlezen option positie in rij " + row + " : " + exc.Message);
            }

            return positions;
        }
        public OptionMarketDataList ReadOptionMarketData(DateTime dtNow, string fileName, ErrorList errors)
        {
            const int ColSecurityName = 1;
            const int ColScenarioName = 2;
            const int ColDividendYield = 8;
            const int ColUnderlyingPrice = 12;
            const int RowStart = 2;

            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, "Options", "A1");
            OptionMarketDataList list = new OptionMarketDataList();
            for (int row = RowStart; row <= values.GetUpperBound(DimensionRow); row++)
            {
                try
                {
                    if (values[row, ColSecurityName] == null)
                        continue;

                    OptionMarketData entry = new OptionMarketData();
                    entry.ScenarioName = values[row, ColScenarioName].ToString().ToLower();
                    entry.DividendYield = (double)values[row, ColDividendYield];
                    entry.UnderlyingPrice = (double)values[row, ColUnderlyingPrice];

                    string name = values[row, ColSecurityName].ToString().ToUpper();
                    list.Add(name, entry);
                    System.Windows.Forms.Application.DoEvents();
                }
                catch (Exception exc)
                {
                    errors.AddError("Fout tijdens inlezen marktdata option in rij " + row + " : " + exc.Message);
                }
            }

            return list;
        }
        // Fixed Income OLD format:
        public PositionList ReadFixedIncomePositions(DateTime dtNow, string fileName, ScenarioList scenarios, ErrorList errors)
        {
            // Defne the base curves:
            CurveList scenarioZeroCurves = new CurveList();
            Curve zeroCurve, scenarioZeroCurve;
            string ccy;
            TotalRisk.Utilities.Scenario baseScenario = scenarios.getScenarioFairValue();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
            {
                ccy = scenarioCurve.m_sName;
                scenarioZeroCurve = scenarioCurve.m_Curve;
                scenarioZeroCurves.Add(ccy.ToUpper(), scenarioZeroCurve);
            }
            // Load Positions:
            const int RowHeader = 1;
            const int RowStart = 2;

            bool invalidDateReported = false;
            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadUsedRangeValues(fileName, "", "A1");
            Dictionary<string, int> colNames = HeaderNamesColumns(values);
            PositionList positions = new PositionList();
            for (int row = RowStart; row <= values.GetUpperBound(DimensionRow); row++)
            {
                try
                {
                    if (values[row, 1] == null)
                    {
                        continue;
                    }
                    string cashflowType = ReadFieldAsString(values, row, 0, colNames, "Type");
                    if (cashflowType.ToLower() != "cash flow")
                    {// first line of the next instrument is "Cash Flow" line
                        continue;
                    }
                    DateTime? dtReport = ReadFieldAsDateTime(values, row, 0, colNames, "Rapportage datum");
                    if (dtReport != dtNow)
                    {
                        string message = "RapportageDatum van kasstroom ongelijk aan rapportage datum in ScenarioTool voor positie in regel " + row.ToString();
                        if (!invalidDateReported)
                        {
                            if (MessageBox.Show(message + ". Bestand alsnog verwerken?", "Ongeldige rapportagedatum", MessageBoxButtons.OKCancel) == DialogResult.OK)
                            {
                                errors.AddWarning("Ongeldige rapportagedatums in bestand. Gebruiker heeft melding genegeerd");
                            }
                            else
                            {
                                errors.AddError("Ongeldige rapportagedatums in bestand. Gebruiker heeft verwerking gestopt");
                                return positions;
                            }

                            invalidDateReported = true;
                        }
                        errors.AddWarning(message);
                    }

                    DateTime? dtMaturity = ReadFieldAsDateTime(values, row, 0, colNames, "Maturity date");

                    CashflowSchedule sched = new CashflowSchedule();
                    int col = HeaderNameId(colNames, "Clean Value PC") + 1;
                    while ((col <= values.GetUpperBound(DimensionCol)) && (values[RowHeader, col] != null))
                    {
                        DateTime? dtBucket = ReadFieldAsDateTime(values, RowHeader, col);
                        int yearBucket = ReadFieldAsInt(values, RowHeader, col);
                        DateTime dtCashflow = DateTimeExtensions.New(yearBucket, dtMaturity.Value.Month, dtMaturity.Value.Day).Value;
                        if (values[row, col] != null) // "Cash Flow" line
                        {
                            Cashflow cf = new Cashflow(dtCashflow, (double)values[row, col]);
                            sched.Add(cf);
                        }
                        col++;
                    }
                    Position position = new Position();

                    if (HeaderNameExists(colNames, "Scope"))
                    {
                        position.m_sScope3 = ReadFieldAsString(values, row, 0, colNames, "Scope");
                    }
                    if (HeaderNameExists(colNames, "Portfolio"))
                    {
                        position.m_sPortfolioId = ReadFieldAsString(values, row, 0, colNames, "Portfolio");
                    }
                    if (HeaderNameExists(colNames, "CIC_LL"))
                    {
                        position.m_sCIC_LL = ReadFieldAsString(values, row, 0, colNames, "CIC_LL");
                        position.m_sCIC_SCR = position.m_sCIC_LL;
                        position.m_sCIC = position.m_sCIC_LL;
                    }
                    if (HeaderNameExists(colNames, "BalanceType"))
                    {
                        position.m_sBalanceType = ReadFieldAsString(values, row, 0, colNames, "BalanceType");
                    }
                    if (HeaderNameExists(colNames, "Group"))
                    {
                        position.m_sGroup = ReadFieldAsString(values, row, 0, colNames, "Group");
                    }
                    if (HeaderNameExists(colNames, "Sub Group"))
                    {
                        position.m_sRiskClass = ReadFieldAsString(values, row, 0, colNames, "Sub Group");
                    }
                    string securityType = ReadFieldAsString(values, row, 0, colNames, "Security Type");
                    string securityId = ReadFieldAsString(values, row, 0, colNames, "Security ID");
                    string securityName = ReadFieldAsString(values, row, 0, colNames, "Security Name");
                    string interestType = ReadFieldAsString(values, row, 0, colNames, "Interest Type");
                    string leg = ReadFieldAsString(values, row, 0, colNames, "Legno");
                    ccy = ReadFieldAsString(values, row, 0, colNames, "Quotation Currency");
                    bool noValue = ("" == ReadFieldAsString(values, row, 0, colNames, "Clean Value PC"));
                    noValue &= ("" == ReadFieldAsString(values, row, 0, colNames, "Accrued interest PC"));
                    double cleanValue = ReadFieldAsDouble(values, row, 0, colNames, "Clean Value PC") + ReadFieldAsDouble(values, row, 0, colNames, "Accrued interest PC");
                    double volume = ReadFieldAsDouble(values, row, 0, colNames, "Nominal Value");
                    double fxRate = ReadFieldAsDouble(values, row, 0, colNames, "FX rate");
                    double coupon = ReadFieldAsDouble(values, row, 0, colNames, "Coupon rate");
                    double collateral = 0;
                    if (HeaderNameExists(colNames, "Collateral coverage percent"))
                    {
                        collateral = ReadFieldAsDouble(values, row, 0, colNames, "Collateral coverage percent");
                    }
                    position.m_sRow = row.ToString();
                    position.m_sSecurityType_LL = securityType;
                    position.m_bICO = ReadFieldAsBool(values, row, 0, colNames, "Ico Code");
                    position.m_sSecurityID_LL = securityId;
                    position.m_sSecurityName_LL = securityName;
                    position.m_sPositionId = position.m_sSecurityID_LL;
                    position.m_sLegId = leg;
                    position.m_sCurrency = ccy;
                    position.m_fFxRate = fxRate;
                    position.m_fVolume = volume;
                    position.m_fCollateralCoveragePercentage = collateral;
                    position.m_sRating = ReadFieldAsString(values, row, 0, colNames, "Rating");

                    Instrument_Cashflow instrument = new Instrument_Cashflow(dtMaturity.Value, sched, coupon);

                    instrument.m_sInstrumentType = securityType;
                    instrument.m_sCouponType = interestType;
                    instrument.m_sLegId = leg;

                    string ccyTranslated = Position.TranslateCurrency_Curve(scenarioZeroCurves, position.m_sCurrency);
                    zeroCurve = scenarioZeroCurves[ccyTranslated];
                    instrument.Init(dtNow, zeroCurve, cleanValue, -0.1); // Spread floored at -10%

                    position.m_Instrument = instrument;
                    position.m_sDATA_Source = "OTHERS";
                    positions.AddPosition(position);

                    if (dtReport > instrument.m_MaturityDate)
                    {
                        position.m_bHasMessage = true;
                        position.m_sMessage = " Warning : Instrument is matured. Fair Value wordt 0 voor alle scenarios.";
                        errors.Add("Row = " + position.m_sRow + " ID = " + position.m_sPositionId + position.m_sMessage);
                    }
                    else
                    {
                        if ((instrument.m_fDirtyValue != 0) && (instrument.m_CashFlowSchedule.Count() == 0))
                        {
                            position.m_bHasMessage = true;
                            position.m_sMessage = " Warning : CleanValue is niet nul, maar er zijn geen kasstromen. Fair Value wordt 0 voor alle scenarios.";
                            errors.Add("Row = " + position.m_sRow + " ID = " + position.m_sPositionId + position.m_sMessage);
                        }
                        else if (double.IsNaN(instrument.m_fImpliedSpread))
                        {
                            position.m_bHasMessage = true;
                            position.m_sMessage = " Warning : Spread kan niet worden bepaald. Fair Value wordt gebruikt voor alle scenarios.";
                            errors.Add("Row = " + position.m_sRow + " ID = " + position.m_sPositionId + position.m_sMessage);
                        }
                    }
                    string CIC2 = position.m_sCIC_SCR.Substring(2, 2);
                    if ("34" == CIC2 || // PREF
                                        //                        "82" == CIC2 || // Spaarlos
                        "75" == CIC2    // deposits to cedants 
                        )
                    {
                        position.m_sSelectieIndex_LL = "1000000001110000"; // rate, valuta, spread, concentration
                        position.m_bEEA = true;
                        position.m_sCountryCurrency = position.m_sCurrency;
                        position.m_fModifiedDuration = instrument.m_fDuration;
                        position.m_fSpreadDuration = position.m_fModifiedDuration;
                        position.m_bGovGuarantee = false;
                        if (position.m_sRating == "" || position.m_sRating.Substring(0, 1) == "NR")
                        {//  seven as NR
                            position.m_dIssuerCreditQuality = 7;
                        }
                        else
                        {
                            position.m_dIssuerCreditQuality = Convert.ToInt32(position.m_sRating.Substring(0, 1));
                        }
                    }
                    if (position.m_sSelectieIndex_LL.Substring(10, 1) == "1") // 11th digit for Spread risk
                    {// not used  for this product
                        position.m_SpreadRiskData = new PositionSpreadRiskData();
                        position.m_SpreadRiskData.m_sSelectieIndex_LL = position.m_sSelectieIndex_LL;
                        position.m_SpreadRiskData.m_sCIC_LL = position.m_sCIC_LL;
                        position.m_SpreadRiskData.m_dSecuritisationType = position.m_dSecuritisationType;
                        position.m_SpreadRiskData.m_sSecurityName_LL = position.m_sSecurityName_LL;
                        position.m_SpreadRiskData.m_sInstrumentID_LL = position.m_sSecurityID_LL;
                        position.m_SpreadRiskData.m_sPortfolioID = position.m_sPortfolioId;
                        position.m_SpreadRiskData.m_sCountryCurrency = position.m_sCountryCurrency;
                        position.m_SpreadRiskData.m_sCurrency = position.m_sCurrency;
                        position.m_SpreadRiskData.m_bEEA = position.m_bEEA;
                        position.m_SpreadRiskData.m_dSecurityCreditQuality = position.m_dSecurityCreditQuality;
                        position.m_SpreadRiskData.m_bGovGuarantee = position.m_bGovGuarantee;
                        position.m_SpreadRiskData.m_fCollateral = position.m_fCollateralCoveragePercentage; // in %
                        position.m_SpreadRiskData.m_fModifiedDuration = position.m_fModifiedDuration;
                        position.m_SpreadRiskData.m_fSpreadDuration = position.m_fSpreadDuration;
                        SCRSpreadModel.WhichProduct(position.m_SpreadRiskData);
                    }

                    System.Windows.Forms.Application.DoEvents();
                }
                catch (Exception exc)
                {
                    errors.AddError("Fout tijdens inlezen vastrentend positie in regel " + row.ToString() + "\n" + exc.Message);
                }
            }

            return positions;
        }
        // Fixed Income IMW format:
        public class CFixedIncomeCashFlowData
        {
            public bool m_bFinished;

            public string m_sUniqSecurityId;
            public int m_dNumberOfCashFlows;
            public int m_dRow;
            public CashflowSchedule m_CashFlowSched;
            public DateTime m_dtReport;
            public DateTime m_dtMaturity;
            public int m_dIndexMaturityCashFlow;
            public string m_sPortfolio;
            public string m_sSecurityType;
            public string m_sSecurityID_LL;
            public string m_sSecurityName;
            public string m_sInterestType;
            public string m_sLeg;
            public string m_sCurrency;

            public double m_fDirtyValue;
            public double m_fVolume;
            public double m_fFxRate;
            public double m_fCoupon;

            public CFixedIncomeCashFlowData(string ID)
            {
                m_bFinished = false;

                m_dIndexMaturityCashFlow = -1;
                m_sUniqSecurityId = ID;
                m_dNumberOfCashFlows = 0;
                m_dRow = 0;

                m_sPortfolio = "";
                m_sSecurityType = "";
                m_sSecurityID_LL = "";
                m_sSecurityName = "";
                m_sInterestType = "";
                m_sLeg = "";
                m_sCurrency = "";

                m_fDirtyValue = 0;

                m_fVolume = 0;
                m_fFxRate = 0;
                m_fCoupon = 0;

                m_CashFlowSched = new CashflowSchedule();
            }
            public int addCashFlow(Cashflow cf)
            {
                m_CashFlowSched.Add(cf);
                m_dNumberOfCashFlows++;
                return m_dNumberOfCashFlows;
            }
            public void finishObject()
            {
                m_dNumberOfCashFlows = m_CashFlowSched.Count();
                if (m_dNumberOfCashFlows > 0)
                {
                    m_dtMaturity = m_CashFlowSched.getCashFlow(0).m_Date;
                    m_dIndexMaturityCashFlow = 0;
                }
                else
                {
                    m_dtMaturity = m_dtReport;
                }
                DateTime temp;
                for (int idx = 1; idx < m_dNumberOfCashFlows; idx++)
                {
                    temp = m_CashFlowSched.getCashFlow(idx).m_Date;
                    if (DateTime.Compare(m_dtMaturity, temp) <= 0)
                    {
                        m_dtMaturity = temp;
                        m_dIndexMaturityCashFlow = idx;
                    }
                }
                m_bFinished = true;
            }

        }
        public Dictionary<string, CFixedIncomeCashFlowData> ReadCashFlowFile_IMW(DateTime dtNow,
            string CashFlow_FileName, Dictionary<string, List<string>> SecurityTypesList, CashFlowType typeCF, ErrorList errors)
        {
            string sCashFlowType = "";
            if (CashFlowType.RiskNeutral == typeCF)
            {
                sCashFlowType = "Cash Flow (risk neutral)";
            }
            else if (CashFlowType.RiskRente == typeCF)
            {
                sCashFlowType = "Cash Flow (rente typisch)";
            }
            else
            {
                return null;
            }
            const int RowStart = 2;
            bool invalidDateReported = false;
            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(CashFlow_FileName, "Kasstromen", "A1");
            Dictionary<string, int> colNames = HeaderNamesColumns(values);
            Dictionary<string, CFixedIncomeCashFlowData> FixedIncomeDataList = new Dictionary<string, CFixedIncomeCashFlowData>();
            int row;
            for (row = RowStart; row <= values.GetUpperBound(DimensionRow); row++)
            {
                //                if (115228 == row)
                //                {
                //                    row += 0;
                //                }
                try
                {
                    if (values[row, 1] == null)
                    {
                        continue;
                    }
                    string test = ReadFieldAsString(values, row, 0, colNames, "Time");
                    if ("" == test)
                    {
                        continue;
                    }
                    test = ReadFieldAsString(values, row, 0, colNames, sCashFlowType);
                    if ("" == test)
                    {
                        continue;
                    }
                    string sSecurityType = ReadFieldAsString(values, row, 0, colNames, "Security Type");
                    if (!SecurityTypesList.ContainsKey(sSecurityType))
                    {
                        continue;
                    }
                    string sLeg = ReadFieldAsString(values, row, 0, colNames, "Leg").Trim();
                    if (!SecurityTypesList[sSecurityType].Contains(sLeg))
                    {
                        continue;
                    }
                    string sPortfolio = ReadFieldAsString(values, row, 0, colNames, "Portfolio").Trim();
                    string sSecurityID_LL = ReadFieldAsString(values, row, 0, colNames, "ID").Trim();
                    string UniqSecurityID = sPortfolio + "_" + sSecurityID_LL + "_" + sLeg;

                    if ("SW812156" == sSecurityID_LL)
                    {
                        sSecurityID_LL += "";
                    }

                    DateTime? dtReport = ReadFieldAsDateTime(values, row, 0, colNames, "Reporting date");
                    if (dtReport != dtNow)
                    {
                        string message = "RapportageDatum van kasstroom ongelijk aan rapportage datum in ScenarioTool voor positie in regel " + row.ToString();
                        if (!invalidDateReported)
                        {
                            if (MessageBox.Show(message + ". Bestand alsnog verwerken?", "Ongeldige rapportagedatum", MessageBoxButtons.OKCancel) == DialogResult.OK)
                            {
                                errors.AddWarning("Ongeldige rapportagedatums in bestand. Gebruiker heeft melding genegeerd");
                            }
                            else
                            {
                                errors.AddError("Ongeldige rapportagedatums in bestand. Gebruiker heeft verwerking gestopt");
                                return new Dictionary<string, CFixedIncomeCashFlowData>();
                            }
                            invalidDateReported = true;
                        }
                        errors.AddWarning(message);
                    }
                    bool FoundInstrumentID = FixedIncomeDataList.ContainsKey(UniqSecurityID) ? true : false;
                    CFixedIncomeCashFlowData FixedIncomeObj;
                    if (!FoundInstrumentID)
                    {
                        FixedIncomeObj = new CFixedIncomeCashFlowData(UniqSecurityID);
                        FixedIncomeDataList.Add(UniqSecurityID, FixedIncomeObj);
                        // static data
                        FixedIncomeObj.m_dRow = row;
                        FixedIncomeObj.m_dtReport = dtReport.Value;
                        FixedIncomeObj.m_sPortfolio = sPortfolio;
                        FixedIncomeObj.m_sSecurityType = sSecurityType;
                        FixedIncomeObj.m_sSecurityID_LL = sSecurityID_LL;
                        FixedIncomeObj.m_sSecurityName = ReadFieldAsString(values, row, 0, colNames, "Security Name");  // ?
                        FixedIncomeObj.m_sLeg = sLeg;
                        FixedIncomeObj.m_sInterestType = ReadFieldAsString(values, row, 0, colNames, "Interest type").ToUpper(); // fixed
                        FixedIncomeObj.m_sCurrency = ReadFieldAsString(values, row, 0, colNames, "Quotation Currency");

                        //                        FixedIncomeObj.m_fDirtyValue = ReadFieldAsDouble(values, row, 0, colNames, "MW");
                        FixedIncomeObj.m_fVolume = ReadFieldAsDouble(values, row, 0, colNames, "Nomal Basis");
                        FixedIncomeObj.m_fFxRate = ReadFieldAsDouble(values, row, 0, colNames, "FX Rate QC PC");
                        FixedIncomeObj.m_fCoupon = ReadFieldAsDouble(values, row, 0, colNames, "Couponrate");
                    }
                    FixedIncomeObj = FixedIncomeDataList[UniqSecurityID];
                    double cashFlowValue = ReadFieldAsDouble(values, row, 0, colNames, sCashFlowType);
                    DateTime? cashFlowDate = ReadFieldAsDateTime(values, row, 0, colNames, "Time");
                    Cashflow cf = new Cashflow(cashFlowDate.Value, cashFlowValue * FixedIncomeObj.m_fFxRate); // cash flow in local currency
                    FixedIncomeObj.addCashFlow(cf);
                }
                catch (Exception exc)
                {
                    errors.AddError("Fout tijdens inlezen vastrentend positie in regel " + row.ToString() + "\n" + exc.Message);
                }
            }
            foreach (CFixedIncomeCashFlowData CPositionData in FixedIncomeDataList.Values)
            {
                CPositionData.finishObject();
            }
            return FixedIncomeDataList;
        }
        public PositionList ReadBondForwardPositions_IMW(DateTime dtNow, string IMW_FileName, string CashFlow_FileName, ScenarioList scenarios, ErrorList errors)
        {
            // Instruments only:
            Dictionary<string, List<string>> SecurityTypesList = new Dictionary<string, List<string>>();
            SecurityTypesList.Add("BOND FW", new List<string>());
            SecurityTypesList["BOND FW"].Add("1");
            SecurityTypesList.Add("BOND FUT", new List<string>());
            SecurityTypesList["BOND FUT"].Add("1");
            // Defne the base curves:
            CurveList scenarioZeroCurves = new CurveList();
            Curve zeroCurve, scenarioZeroCurve;
            string ccy;
            TotalRisk.Utilities.Scenario baseScenario = scenarios.getScenarioFairValue();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
            {
                ccy = scenarioCurve.m_sName;
                scenarioZeroCurve = scenarioCurve.m_Curve;
                scenarioZeroCurves.Add(ccy.ToUpper(), scenarioZeroCurve);
            }
            // Load positions from IMW bestand:
            const int RowStart = 2;
            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(IMW_FileName, "Integrale aanlevering IMW maand", "A1");
            Dictionary<string, int> colNames = HeaderNamesColumns(values);
            // Scope : Security Id Ll : (leg, leg data)
            Dictionary<string, Instrument_BondForward_OriginalData_List> IMW_File_Data =
                new Dictionary<string, Instrument_BondForward_OriginalData_List>();
            int row = 0;
            int numberOfSwaps = 0;
            string sErrorMessage;
            try
            {
                for (row = RowStart; row <= values.GetUpperBound(DimensionRow); row++)
                {
                    sErrorMessage = "";
                    string reportCode = ReadFieldAsString(values, row, 0, colNames, "SelectieIndex LL");
                    if (reportCode.Substring(0, 1) != "1") // first digit for Cash flow report
                    {
                        continue;
                    }
                    BondForwardType bondForwardType;
                    string securityType = ReadFieldAsString(values, row, 0, colNames, "Security Type Ll").ToUpper();
                    string cic_ID = ReadFieldAsString(values, row, 0, colNames, "Cic Id Ll").ToUpper();
                    string cic_ID_last2 = cic_ID.Substring(2, 2);
                    if (cic_ID_last2 == "E9" && "BOND FW" == securityType)
                    {
                        bondForwardType = BondForwardType.Forward;
                    }
                    else if (cic_ID_last2 == "A2" && "BOND FUT" == securityType)
                    {
                        bondForwardType = BondForwardType.Futures;
                    }
                    else
                    {
                        continue;
                    }

                    Instrument_BondForward_OriginalData contractData = new Instrument_BondForward_OriginalData();
                    contractData.m_dRow = row;
                    contractData.m_BondForwardType = bondForwardType;
                    double dateValue = ReadFieldAsDouble(values, row, 0, colNames, "Reporting Date");
                    contractData.m_dtReport = getExcelDate_From_DoubleDate(dateValue);
                    contractData.m_sScope3 = ReadFieldAsString(values, row, 0, colNames, "Ecs Cons Ecap asr");
                    contractData.m_sScope3 = ScopeData.getScopeFormated(contractData.m_sScope3);
                    contractData.m_sSelectieIndex_LL = reportCode;
                    //                    contractData.m_fMarketValue = ReadFieldAsDouble(values, row, 0, colNames, "Market Value Eur Ll");
                    contractData.m_fAccruedInterest_LL = ReadFieldAsDouble(values, row, 0, colNames, "Accrued Interest Ll");
                    contractData.m_fCollateral = ReadFieldAsDouble(values, row, 0, colNames, "Coll Coverage Laagste Level");

                    contractData.m_fExpiryDate = ReadFieldAsDouble(values, row, 0, colNames, "Maturity LL");
                    contractData.m_fStrikePrice = ReadFieldAsDouble(values, row, 0, colNames, "Strike Price Laagste Level") / 100;
                    contractData.m_fConversionFactor = ReadFieldAsDouble(values, row, 0, colNames, "Conversion Ratio for CTD Ll");
                    if (contractData.m_fConversionFactor <= 0)
                    {
                        sErrorMessage = " - ERROR - Conversion Factor must be non-zero and positive - field ( " + "Conversion Ratio for CTD Ll" + " )!";
                        throw new Exception(sErrorMessage);
                    }
                    contractData.m_sUnderlyingSecurityCode = ReadFieldAsString(values, row, 0, colNames, "Underlying security");
                    contractData.m_sUnderlyingSecurityName = ReadFieldAsString(values, row, 0, colNames, "Underlying name");
                    contractData.m_fNominal = ReadFieldAsDouble(values, row, 0, colNames, "Balnomval qc");
                    if (BondForwardType.Forward == bondForwardType)
                    {
                        contractData.m_fUnderlyingSecurityValue = ReadFieldAsDouble(values, row, 0, colNames, "Underlying MarketValue"); // clean value.
                        contractData.m_fVariationMargin = ReadFieldAsDouble(values, row, 0, colNames, "Variation Margin PC");
                        contractData.m_fMarketValue = ReadFieldAsDouble(values, row, 0, colNames, "Variation Margin PC");
                    }
                    else if (BondForwardType.Futures == bondForwardType)
                    {
                        contractData.m_fUnderlyingSecurityValue = ReadFieldAsDouble(values, row, 0, colNames, "Future exposure");
                        contractData.m_fVariationMargin = ReadFieldAsDouble(values, row, 0, colNames, "Variation Margin PC");
                        contractData.m_fMarketValue = contractData.m_fUnderlyingSecurityValue;
                    }
                    contractData.m_sPortfolioID = ReadFieldAsString(values, row, 0, colNames, "Portfolio Id");
                    contractData.m_sCIC_ID_LL = cic_ID;
                    contractData.m_sSecurityID_LL = ReadFieldAsString(values, row, 0, colNames, "Security Id Ll");
                    contractData.m_sLeg = ReadFieldAsString(values, row, 0, colNames, "Leg no");

                    contractData.m_sUniquePositionId = contractData.m_sPortfolioID + "_" + contractData.m_sSecurityID_LL + "_1";

                    contractData.m_sSecurityName = ReadFieldAsString(values, row, 0, colNames, "Security Name Ll");
                    contractData.m_sSecurityType = securityType;
                    contractData.m_sCurrency = ReadFieldAsString(values, row, 0, colNames, "Currency Laagste Lt Level").ToUpper().Trim();
                    contractData.m_fFxRate = 1.0 / ReadFieldAsDouble(values, row, 0, colNames, "Fx Rate Qc Pc Laagste Lt Level");
                    contractData.m_sType = ReadFieldAsString(values, row, 0, colNames, "Derivaten Type Ll");

                    contractData.m_bGovGuarantee = ReadFieldAsBool(values, row, 0, colNames, "Gov Guaranteed Laagste Level");
                    contractData.m_bEEA = ReadFieldAsBool(values, row, 0, colNames, "EEA land"); // new OD
                    contractData.m_dSecuritisationType = ReadFieldAsInt(values, row, 0, colNames, "Type 1 or 2 Ll"); // new OD

                    contractData.m_bICO = ReadFieldAsBool(values, row, 0, colNames, "Eliminatie ASR");
                    contractData.m_sNACEcode = ReadFieldAsString(values, row, 0, colNames, "Nace laagste LT level Ll");
                    contractData.m_sSecurityID = ReadFieldAsString(values, row, 0, colNames, "Security Id");
                    contractData.m_sCIC_ID = ReadFieldAsString(values, row, 0, colNames, "Cic Id").ToUpper();
                    contractData.m_sAccount = ReadFieldAsString(values, row, 0, colNames, "RDS-STA account");
                    contractData.m_sAccount_LL = ReadFieldAsString(values, row, 0, colNames, "RDS-STA account LT");
                    contractData.m_sECAP_Category_LL = ReadFieldAsString(values, row, 0, colNames, "ECAP Category Ll");

                    // credit quality steps issuer:
                    string test = ReadFieldAsString(values, row, 0, colNames, "Issuer Cr Quality Step Laagste Level");
                    contractData.m_sIssuerCreditQuality = test;
                    if (test == "" || test.Substring(0, 1) == "NR")
                    {//  seven as NR
                        contractData.m_dIssuerCreditQuality = 7;
                    }
                    else
                    {
                        contractData.m_dIssuerCreditQuality = Convert.ToInt32(test.Substring(0, 1));
                    }
                    // cred quality step
                    test = ReadFieldAsString(values, row, 0, colNames, "Credit Quality Step Ll");
                    contractData.m_sSecurityCreditQuality = test;
                    if (test == "" || test.Substring(0, 1) == "NR")
                    {//  seven as NR
                        contractData.m_dSecurityCreditQuality = 7;
                    }
                    else
                    {
                        contractData.m_dSecurityCreditQuality = Convert.ToInt32(test.Substring(0, 1));
                    }
                    if (contractData.m_dSecurityCreditQuality > 7)
                    {
                        contractData.m_dSecurityCreditQuality = 7;
                    }
                    // credit quality steps Group Vounterparty:
                    contractData.m_sGroupCounterpartyName = ReadFieldAsString(values, row, 0, colNames, "Groep tegenpartij naam Ll");
                    contractData.m_sGroupCounterpartyLEI = ReadFieldAsString(values, row, 0, colNames, "Groep tegenpartij LEI Ll");
                    test = ReadFieldAsString(values, row, 0, colNames, "Groep tegenpartij Credit Quality Step Ll");
                    contractData.m_sGroupCounterpartyCQS = test;
                    if (test == "" || test.Substring(0, 1) == "NR")
                    {//  seven as NR
                        contractData.m_dGroupCounterpartyCQS = 7;
                    }
                    else
                    {
                        contractData.m_dGroupCounterpartyCQS = Convert.ToInt32(test.Substring(0, 1));
                    }
                    //Call Date
                    DateTime? callDate = getExcelDate_From_DoubleDate(contractData.m_fExpiryDate);
                    test = ReadFieldAsString(values, row, 0, colNames, "Maturity Call LL");
                    if ("" == test)
                    {
                        contractData.m_bCallable = false;
                    }
                    else
                    {
                        dateValue = ReadFieldAsDouble(values, row, 0, colNames, "Maturity Call LL");
                        callDate = getExcelDate_From_DoubleDate(dateValue);
                        if (callDate > contractData.m_dtReport)
                        {
                            contractData.m_bCallable = true;
                        }
                        else
                        {
                            contractData.m_bCallable = false;
                        }
                    }

                    // Interest rate duration:
                    test = ReadFieldAsString(values, row, 0, colNames, "Mod Dur Laagste Lt Level");
                    if ("" == test)
                    {
                        contractData.m_fModifiedDurationOrig = 0;
                    }
                    else
                    {
                        contractData.m_fModifiedDurationOrig = ReadFieldAsDouble(values, row, 0, colNames, "Mod Dur Laagste Lt Level");
                    }
                    // Spread duration:
                    test = ReadFieldAsString(values, row, 0, colNames, "Mod Spread Dur Laagste Level");
                    if ("" == test)
                    {
                        contractData.m_fSpreadDurationOrig = 0;
                    }
                    else
                    {
                        contractData.m_fSpreadDurationOrig = ReadFieldAsDouble(values, row, 0, colNames, "Mod Spread Dur Laagste Level");
                    }
                    // Checking and Correction of durations:
                    double spreadDuration = contractData.m_fSpreadDurationOrig;
                    double interestRateDuration = contractData.m_fModifiedDurationOrig;
                    if (spreadDuration == 0 && interestRateDuration != 0)
                    {
                        spreadDuration = interestRateDuration;
                    }
                    contractData.m_fSpreadDurationCorrected = spreadDuration;
                    contractData.m_fModifiedDurationCorrected = interestRateDuration;
                    // DNB Stress test data if they are present:
                    if (HeaderNameExists(colNames, "DNB Government Bonds"))
                    {
                        if (ReadFieldAsBool(values, row, 0, colNames, "DNB Government Bonds"))
                        {
                            contractData.m_DNB_Type = DNB_Bond.Government;
                            contractData.m_sDNB_CountryUnion = ReadFieldAsString(values, row, 0, colNames, "DNB Land code").Trim().ToUpper();
                        }
                    }

                    if (!IMW_File_Data.ContainsKey(contractData.m_sSecurityID_LL))
                    {
                        IMW_File_Data.Add(contractData.m_sSecurityID_LL, new Instrument_BondForward_OriginalData_List());
                        numberOfSwaps++;
                    }
                    IMW_File_Data[contractData.m_sSecurityID_LL].Add(contractData);
                }
            }
            catch (Exception exc)
            {
                errors.AddError("Fout tijdens inlezen BOND FORWARD positie in rij " + row + " : " + exc.Message);
            }

            // Read Cash Flow file:
            Dictionary<string, CFixedIncomeCashFlowData> FixedIncomeDataList = ReadCashFlowFile_IMW(dtNow, CashFlow_FileName,
                SecurityTypesList, CashFlowType.RiskRente, errors);
            if (0 == FixedIncomeDataList.Count)
            {
                return new PositionList();
            }
            // Create the position list:
            PositionList positions = new PositionList();
            foreach (KeyValuePair<string, Instrument_BondForward_OriginalData_List> CSecurityData in IMW_File_Data)
            {
                string m_sSecurityID_LL = CSecurityData.Key;
                Instrument_BondForward_OriginalData_List DataList = CSecurityData.Value;
                foreach (Instrument_BondForward_OriginalData CPositionData in DataList)
                {
                    bool FoundInstrumentID = FixedIncomeDataList.ContainsKey(CPositionData.m_sUniquePositionId) ? true : false;
                    if (!FoundInstrumentID)
                    {
                        string message = "No Cash Flow for Portfolio (" + CPositionData.m_sPortfolioID + ") Security (" + CPositionData.m_sSecurityID_LL + ")  Leg (1)";
                        if (MessageBox.Show(message + ". Bestand alsnog verwerken?", "Ongeldige rapportagedatum", MessageBoxButtons.OKCancel) == DialogResult.OK)
                        {
                            errors.AddWarning(message + " Gebruiker heeft melding genegeerd");
                        }
                        else
                        {
                            errors.AddError(message + " Gebruiker heeft verwerking gestopt");
                            return new PositionList();
                        }
                    }
                    CFixedIncomeCashFlowData FixedIncomeObj = FixedIncomeDataList[CPositionData.m_sUniquePositionId];

                    Position position = new Position();
                    position.m_sDATA_Source = "IMW";
                    position.m_sRow = CPositionData.m_dRow.ToString();
                    position.m_bIsLookThroughPosition = false;

                    position.m_sBalanceType = "Assets";
                    position.m_sGroup = "Bond Derivatives";
                    position.m_sUniquePositionId = CPositionData.m_sUniquePositionId;
                    position.m_sSelectieIndex_LL = CPositionData.m_sSelectieIndex_LL;
                    position.m_sSecurityType_LL = CPositionData.m_sSecurityType;
                    position.m_bICO = CPositionData.m_bICO;
                    position.m_sScope3 = CPositionData.m_sScope3;
                    position.m_sSecurityID_LL = CPositionData.m_sSecurityID_LL;
                    position.m_sSecurityName_LL = CPositionData.m_sSecurityName;
                    position.m_sPositionId = position.m_sSecurityID_LL;
                    position.m_sPortfolioId = CPositionData.m_sPortfolioID;

                    position.m_sLegId = CPositionData.m_sLeg;
                    position.m_sCurrency = CPositionData.m_sCurrency;
                    position.m_sCountryCurrency = CPositionData.m_sCurrency;
                    position.m_fFxRate = CPositionData.m_fFxRate;
                    position.m_fVolume = CPositionData.m_fNominal;
                    position.m_sRiskClass = CPositionData.m_sNACEcode;
                    position.m_sCIC_LL = CPositionData.m_sCIC_ID_LL;
                    position.m_sCIC_SCR = position.m_sCIC_LL;
                    position.m_bGovGuarantee = CPositionData.m_bGovGuarantee;
                    position.m_bEEA = CPositionData.m_bEEA;
                    position.m_fCollateralCoveragePercentage = CPositionData.m_fCollateral;
                    position.m_dSecuritisationType = CPositionData.m_dSecuritisationType;
                    position.m_sSecurityCreditQuality = CPositionData.m_sSecurityCreditQuality;
                    position.m_dSecurityCreditQuality = CPositionData.m_dSecurityCreditQuality;
                    position.m_sIssuerCreditQuality = CPositionData.m_sSecurityCreditQuality;
                    position.m_dIssuerCreditQuality = CPositionData.m_dIssuerCreditQuality;
                    position.m_fModifiedDuration = CPositionData.m_fModifiedDurationCorrected;
                    position.m_fSpreadDuration = CPositionData.m_fSpreadDurationCorrected;
                    DateTime Maturity = getExcelDate_From_DoubleDate(CPositionData.m_fExpiryDate).Value;

                    Instrument_BondForward instrument = new Instrument_BondForward(Maturity, CPositionData, FixedIncomeObj);

                    string ccyTranslated = Position.TranslateCurrency_Curve(scenarioZeroCurves, position.m_sCurrency);
                    zeroCurve = scenarioZeroCurves[ccyTranslated];

                    instrument.Init(dtNow, zeroCurve, CPositionData.m_fUnderlyingSecurityValue, CPositionData.m_fMarketValue); // Spread floored at -10%
                    position.m_Instrument = instrument;

                    position.m_sSecurityID = CPositionData.m_sSecurityID;
                    position.m_sCIC = CPositionData.m_sCIC_ID;
                    position.m_fAccruedInterest_LL = CPositionData.m_fAccruedInterest_LL;
                    position.m_fAccruedDividend_LL = 0;
                    position.m_sAccount = CPositionData.m_sAccount;
                    position.m_sAccount_LL = CPositionData.m_sAccount_LL;
                    position.m_sECAP_Category_LL = CPositionData.m_sECAP_Category_LL;
                    if (position.m_sSelectieIndex_LL.Substring(10, 1) == "1") // 11th digit for Spread risk
                    {// not used  for this product
                        position.m_SpreadRiskData = new PositionSpreadRiskData();
                        position.m_SpreadRiskData.m_sSelectieIndex_LL = position.m_sSelectieIndex_LL;
                        position.m_SpreadRiskData.m_sCIC_LL = position.m_sCIC_LL;
                        position.m_SpreadRiskData.m_dSecuritisationType = position.m_dSecuritisationType;
                        position.m_SpreadRiskData.m_sSecurityName_LL = position.m_sSecurityName_LL;
                        position.m_SpreadRiskData.m_sInstrumentID_LL = position.m_sSecurityID_LL;
                        position.m_SpreadRiskData.m_sPortfolioID = position.m_sPortfolioId;
                        position.m_SpreadRiskData.m_sCountryCurrency = position.m_sCountryCurrency;
                        position.m_SpreadRiskData.m_sCurrency = position.m_sCurrency;
                        position.m_SpreadRiskData.m_bEEA = position.m_bEEA;
                        position.m_SpreadRiskData.m_dSecurityCreditQuality = position.m_dSecurityCreditQuality;
                        position.m_SpreadRiskData.m_bGovGuarantee = position.m_bGovGuarantee;
                        position.m_SpreadRiskData.m_fCollateral = position.m_fCollateralCoveragePercentage; // in %
                        position.m_SpreadRiskData.m_fModifiedDuration = position.m_fModifiedDuration;
                        position.m_SpreadRiskData.m_fSpreadDuration = position.m_fSpreadDuration;
                        SCRSpreadModel.WhichProduct(position.m_SpreadRiskData);
                    }

                    positions.AddPosition(position);

                }
            }
            return positions;
        }
        public PositionList ReadBondPositions_IMW(DateTime dtNow, string sPeriod, string IMW_FileName, string CashFlow_FileName, ScenarioList scenarios, FixedIncomeModels model, ErrorList errors)
        {
            // Single leg Instruments only:
            Dictionary<string, List<string>> SecurityTypesList = new Dictionary<string, List<string>>();
            SecurityTypesList.Add("BOND", new List<string>());
            SecurityTypesList.Add("BOND ZERO", new List<string>());
            SecurityTypesList.Add("BOND FRN", new List<string>());
            SecurityTypesList.Add("INDEX BOND", new List<string>());
            SecurityTypesList.Add("LOAN", new List<string>());
            SecurityTypesList.Add("LOAN ZERO", new List<string>());
            SecurityTypesList.Add("LOAN CLIMB", new List<string>());
            SecurityTypesList.Add("PERPETUAL", new List<string>());
            SecurityTypesList.Add("ABS", new List<string>());
            SecurityTypesList.Add("DEPOSIT", new List<string>());
            SecurityTypesList.Add("PREF", new List<string>());
            SecurityTypesList.Add("LOAN SUB", new List<string>());
            //            SecurityTypes.Add("CALL MONEY");
            // Defne the base curves:
            Curve zeroCurve, scenarioZeroCurve;
            IndexCPI CPIobject;
            string ccy;
            TotalRisk.Utilities.Scenario baseScenario = scenarios.getScenarioFairValue();
            CurveList scenarioZeroCurves = new CurveList();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
            {
                ccy = scenarioCurve.m_sName;
                scenarioZeroCurve = scenarioCurve.m_Curve;
                scenarioZeroCurves.Add(ccy.ToUpper(), scenarioZeroCurve);
            }
            CurveList scenarioZeroInflationCurves = new CurveList();
            SortedList<string, IndexCPI> scenarioIndexCPI_List = new SortedList<string, IndexCPI>();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_InflationCurves)
            {
                ccy = scenarioCurve.m_sName;
                scenarioZeroInflationCurves.Add(ccy.ToUpper(), scenarioCurve.m_Curve);
                CPIobject = new IndexCPI();
                CPIobject.SetInflationInstance(dtNow, 100, scenarioCurve.m_Curve, 100);
                scenarioIndexCPI_List.Add(ccy.ToUpper(), CPIobject);
            }

            // Load positions from IMW bestand:
            const int RowStart = 2;
            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(IMW_FileName, "Integrale aanlevering IMW maand", "A1");
            Dictionary<string, int> colNames = HeaderNamesColumns(values);
            // Scope : Security Id Ll : (leg, leg data)
            Dictionary<string, Instrument_Bond_OriginalData_List> IMW_File_Data =
                new Dictionary<string, Instrument_Bond_OriginalData_List>();
            int row = 0;
            int numberOfSwaps = 0;
            string sErrorMessage;
            try
            {
                for (row = RowStart; row <= values.GetUpperBound(DimensionRow); row++)
                {
                    sErrorMessage = "";
                    string reportCode = ReadFieldAsString(values, row, 0, colNames, "SelectieIndex LL");
                    bool bLookthroughData = false;
                    if (reportCode.Substring(1, 1) == "1") // the second digit stays for lookthrough data
                    {
                        bLookthroughData = true;
                    }
                    if (reportCode.Substring(0, 1) != "1") // the first digit stays for Interest rate risk in the Cash Flow
                    {
                        continue;
                    }
                    string sSecurity_ID = ReadFieldAsString(values, row, 0, colNames, "Security Id");
                    //                    if (sSecurity_ID == "GRR000000010")
                    //                    {
                    //                        sSecurity_ID += "";
                    //                    }
                    BondType bondType = BondType.CASH_FLOW;
                    string securityType_LL = ReadFieldAsString(values, row, 0, colNames, "Security Type Ll").ToUpper();
                    string cic_ID_LL = ReadFieldAsString(values, row, 0, colNames, "Cic Id Ll").ToUpper();
                    string CIC34_LL = cic_ID_LL.Substring(2, 2);
                    string cic_ID = ReadFieldAsString(values, row, 0, colNames, "Cic Id").ToUpper();
                    string CIC34 = cic_ID.Substring(2, 2);
                    if (bLookthroughData)
                    {
                        bondType = BondType.CASH_FLOW;
                    }
                    else if (SecurityTypesList.ContainsKey(securityType_LL))
                    {
                        bondType = BondType.CASH_FLOW;
                        if (securityType_LL.Equals("INDEX BOND"))
                        {
                            bondType = BondType.INFLATION;
                        }
                        //                        continue; // test look through data only
                    }
                    else
                    {
                        continue;
                    }

                    Instrument_Bond_OriginalData contractData = new Instrument_Bond_OriginalData();
                    contractData.m_dRow = row;
                    contractData.m_BondType = bondType;
                    contractData.m_bLookthroughData = bLookthroughData;
                    double dateValue = ReadFieldAsDouble(values, row, 0, colNames, "Reporting Date");
                    contractData.m_dtReport = getExcelDate_From_DoubleDate(dateValue);
                    if (null == contractData.m_dtReport)
                    {
                        sErrorMessage = " - ERROR - Date format is wrong in field ( " + "Reporting Date" + " )!";
                        throw new Exception(sErrorMessage);
                    }
                    contractData.m_sScope3 = ReadFieldAsString(values, row, 0, colNames, "Ecs Cons Ecap asr");
                    contractData.m_sScope3 = ScopeData.getScopeFormated(contractData.m_sScope3);
                    contractData.m_sSelectieIndex_LL = reportCode;
                    contractData.m_fMarketValue = ReadFieldAsDouble(values, row, 0, colNames, "Market Value Eur Ll");
                    if ("22" == CIC34_LL)
                    {// Convertible Bonds
                        double fMarketValue_EmbededOption = ReadFieldAsDouble(values, row, 0, colNames, "Convertible Optie Waarde Ll");
                        contractData.m_fMarketValue -= fMarketValue_EmbededOption;
                    }
                    contractData.m_fAccruedInterest_LL = ReadFieldAsDouble(values, row, 0, colNames, "Accrued Interest Ll");
                    contractData.m_fCollateral = ReadFieldAsDouble(values, row, 0, colNames, "Coll Coverage Laagste Level");

                    contractData.m_fCoupon = ReadFieldAsDouble(values, row, 0, colNames, "Coupon Perc Laagste Lt Level") / 100;
                    contractData.m_sCouponType = ReadFieldAsString(values, row, 0, colNames, "Coupon Type Laagste Lt Level");
                    contractData.m_dCouponFrequency = ReadFieldAsInt(values, row, 0, colNames, "Coupon Frequency Laagste Level");
                    contractData.m_fCouponReferenceRate = ReadFieldAsDouble(values, row, 0, colNames, "Coupon Reference Rate Ll") / 100;
                    contractData.m_fCouponSpread = ReadFieldAsDouble(values, row, 0, colNames, "Coupon Spread Ll") / 100;
                    //                    dateValue = ReadFieldAsDouble(values, row, 0, colNames, "First Coupon Data Ll");
                    //                    contractData.m_dtFirstCouponDate = getExcelDateFromDoubleDate(dateValue);

                    contractData.m_sPortfolioID = ReadFieldAsString(values, row, 0, colNames, "Portfolio Id");
                    contractData.m_sCIC_ID_LL = cic_ID_LL;
                    contractData.m_sCIC_ID = cic_ID;
                    contractData.m_sSecurityID = ReadFieldAsString(values, row, 0, colNames, "Security Id");
                    contractData.m_sSecurityName = ReadFieldAsString(values, row, 0, colNames, "Security Name");
                    contractData.m_sSecurityID_LL = ReadFieldAsString(values, row, 0, colNames, "Security Id Ll");
                    contractData.m_sLeg = ReadFieldAsString(values, row, 0, colNames, "Leg no").Trim();

                    if (contractData.m_bLookthroughData)
                    {
                        contractData.m_sUniquePositionId = "LT_" + contractData.m_sPortfolioID + "_" + contractData.m_sSecurityID + "_" + contractData.m_sSecurityID_LL + "_" + contractData.m_sLeg;
                    }
                    else
                    {
                        contractData.m_sUniquePositionId = contractData.m_sPortfolioID + "_" + contractData.m_sSecurityID_LL + "_" + contractData.m_sLeg;
                    }
                    contractData.m_sSecurityName_LL = ReadFieldAsString(values, row, 0, colNames, "Security Name Ll");
                    contractData.m_sSecurityType_LL = securityType_LL;
                    contractData.m_sCurrencyCountry = ReadFieldAsString(values, row, 0, colNames, "Country Currency Ll").ToUpper().Trim();
                    contractData.m_sCurrency = ReadFieldAsString(values, row, 0, colNames, "Currency Laagste Lt Level").ToUpper().Trim();

                    double FXRate = 1;
                    string test = ReadFieldAsString(values, row, 0, colNames, "Fx Rate Qc Pc Laagste Lt Level");
                    if (test == "")
                    {
                        FXRate = 1;
                    }
                    else
                    {
                        FXRate = ReadFieldAsDouble(values, row, 0, colNames, "Fx Rate Qc Pc Laagste Lt Level");
                    }
                    if (FXRate <= 0)
                    {
                        FXRate = 1;
                    }
                    contractData.m_fFxRate = 1.0 / FXRate; // Foreign Curr in EUR
                    contractData.m_fNominal = ReadFieldAsDouble(values, row, 0, colNames, "Balnomval qc");
                    if (contractData.m_bLookthroughData)
                    {
                        //                        if (0 == contractData.m_fNominal)
                        {
                            contractData.m_fNominal = contractData.m_fMarketValue / contractData.m_fFxRate;
                        }
                    }
                    contractData.m_sType = ReadFieldAsString(values, row, 0, colNames, "Derivaten Type Ll");

                    contractData.m_bGovGuarantee = ReadFieldAsBool(values, row, 0, colNames, "Gov Guaranteed Laagste Level");
                    contractData.m_bEEA = ReadFieldAsBool(values, row, 0, colNames, "EEA land"); // new OD
                    contractData.m_dSecuritisationType = ReadFieldAsInt(values, row, 0, colNames, "Type 1 or 2 Ll");

                    contractData.m_bICO = ReadFieldAsBool(values, row, 0, colNames, "Eliminatie ASR");
                    contractData.m_sNACEcode = ReadFieldAsString(values, row, 0, colNames, "Nace laagste LT level Ll");
                    contractData.m_sSecurityID = ReadFieldAsString(values, row, 0, colNames, "Security Id");
                    contractData.m_sCIC_ID = cic_ID;
                    contractData.m_sAccount = ReadFieldAsString(values, row, 0, colNames, "RDS-STA account");
                    contractData.m_sAccount_LL = ReadFieldAsString(values, row, 0, colNames, "RDS-STA account LT");
                    contractData.m_sECAP_Category_LL = ReadFieldAsString(values, row, 0, colNames, "ECAP Category Ll");
                    contractData.m_sPortfolioPurpose = ReadFieldAsString(values, row, 0, colNames, "Portfolio-purpose").Trim().ToUpper();

                    // credit quality steps issuer:
                    test = ReadFieldAsString(values, row, 0, colNames, "Issuer Cr Quality Step Laagste Level");
                    contractData.m_sIssuerCreditQuality = test;
                    if (test == "" || test.Substring(0, 1) == "NR")
                    {//  seven as NR
                        contractData.m_dIssuerCreditQuality = 7;
                    }
                    else
                    {
                        contractData.m_dIssuerCreditQuality = Convert.ToInt32(test.Substring(0, 1));
                    }
                    // cred quality step
                    test = ReadFieldAsString(values, row, 0, colNames, "Credit Quality Step Ll");
                    contractData.m_sSecurityCreditQuality = test;
                    if (test == "" || test.Substring(0, 1) == "NR")
                    {//  seven as NR
                        contractData.m_dSecurityCreditQuality = 7;
                    }
                    else
                    {
                        contractData.m_dSecurityCreditQuality = Convert.ToInt32(test.Substring(0, 1));
                    }
                    if (contractData.m_dSecurityCreditQuality > 7)
                    {
                        contractData.m_dSecurityCreditQuality = 7;
                    }
                    // credit quality steps Group Vounterparty:
                    contractData.m_sGroupCounterpartyName = ReadFieldAsString(values, row, 0, colNames, "Groep tegenpartij naam Ll");
                    contractData.m_sGroupCounterpartyLEI = ReadFieldAsString(values, row, 0, colNames, "Groep tegenpartij LEI Ll");
                    test = ReadFieldAsString(values, row, 0, colNames, "Groep tegenpartij Credit Quality Step Ll");
                    contractData.m_sGroupCounterpartyCQS = test;
                    if (test == "" || test.Substring(0, 1) == "NR")
                    {//  seven as NR
                        contractData.m_dGroupCounterpartyCQS = 7;
                    }
                    else
                    {
                        contractData.m_dGroupCounterpartyCQS = Convert.ToInt32(test.Substring(0, 1));
                    }
                    //Call Date
                    test = ReadFieldAsString(values, row, 0, colNames, "Maturity Call LL");
                    if ("" == test)
                    {
                        contractData.m_bCallable = false;
                        contractData.m_fCallDate = 0;
                    }
                    else
                    {
                        dateValue = ReadFieldAsDouble(values, row, 0, colNames, "Maturity Call LL");
                        contractData.m_fCallDate = dateValue;
                        DateTime? callDate = getExcelDate_From_DoubleDate(dateValue);
                        if (callDate > contractData.m_dtReport)
                        {
                            contractData.m_bCallable = true;
                        }
                        else
                        {
                            contractData.m_bCallable = false;
                        }
                    }
                    // Interest rate duration:
                    test = ReadFieldAsString(values, row, 0, colNames, "Mod Dur Laagste Lt Level");
                    if ("" == test)
                    {
                        contractData.m_fModifiedDurationOrig = 0;
                    }
                    else
                    {
                        contractData.m_fModifiedDurationOrig = ReadFieldAsDouble(values, row, 0, colNames, "Mod Dur Laagste Lt Level");
                    }
                    // Spread duration:
                    test = ReadFieldAsString(values, row, 0, colNames, "Mod Spread Dur Laagste Level");
                    if ("" == test)
                    {
                        contractData.m_fSpreadDurationOrig = 0;
                    }
                    else
                    {
                        contractData.m_fSpreadDurationOrig = ReadFieldAsDouble(values, row, 0, colNames, "Mod Spread Dur Laagste Level");
                    }
                    // Checking and Correction of durations:
                    double spreadDuration = contractData.m_fSpreadDurationOrig;
                    double interestRateDuration = contractData.m_fModifiedDurationOrig;
                    if (spreadDuration == 0 && interestRateDuration != 0)
                    {
                        spreadDuration = interestRateDuration;
                    }
                    contractData.m_fSpreadDurationCorrected = spreadDuration;
                    contractData.m_fModifiedDurationCorrected = interestRateDuration;
                    // Maturity:
                    test = ReadFieldAsString(values, row, 0, colNames, "Maturity LL");
                    if ("" == test)
                    {
                        if (contractData.m_bCallable)
                        {
                            contractData.m_fExpiryDate = contractData.m_fCallDate;
                        }
                        else if (contractData.m_bLookthroughData)
                        {
                            DateTime? maturity = contractData.m_dtReport;
                            if (contractData.m_fModifiedDurationCorrected > 0)
                            {
                                double months = contractData.m_fModifiedDurationCorrected * 12;
                                months = Math.Ceiling(months);
                                maturity = maturity.Value.AddMonths((int)months);
                            }
                            contractData.m_fExpiryDate = getDoubleDate_From_ExcelDate(maturity);
                        }
                        else
                        {
                            contractData.m_fExpiryDate = 0;
                        }
                    }
                    else
                    {
                        contractData.m_fExpiryDate = ReadFieldAsDouble(values, row, 0, colNames, "Maturity LL");
                    }
                    // DNB Stress test data if they are present:
                    if (HeaderNameExists(colNames, "DNB Government Bonds"))
                    {
                        if (ReadFieldAsBool(values, row, 0, colNames, "DNB Government Bonds"))
                        {
                            contractData.m_DNB_Type = DNB_Bond.Government;
                            contractData.m_sDNB_CountryUnion = ReadFieldAsString(values, row, 0, colNames, "DNB Land code").Trim().ToUpper();
                        }
                        else if (ReadFieldAsBool(values, row, 0, colNames, "DNB Corporate Bonds"))
                        {
                            contractData.m_DNB_Type = DNB_Bond.Corporate;
                            contractData.m_sDNB_CountryUnion = ReadFieldAsString(values, row, 0, colNames, "DNB CorB Country/Union").Trim().ToUpper();
                            test = ReadFieldAsString(values, row, 0, colNames, "DNB Financial / non-Financial").Trim().ToUpper();
                            if (test == "F")
                            {
                                contractData.m_bDNB_Financial = true;
                            }
                            contractData.m_sDNB_Rating = ReadFieldAsString(values, row, 0, colNames, "DNB CorB Rating").Trim().ToUpper();
                        }
                        else if (ReadFieldAsBool(values, row, 0, colNames, "DNB Covered Bonds"))
                        {
                            contractData.m_DNB_Type = DNB_Bond.Covered;
                            contractData.m_sDNB_CountryUnion = ReadFieldAsString(values, row, 0, colNames, "DNB CVB Country/Union").Trim().ToUpper();
                            contractData.m_sDNB_Rating = ReadFieldAsString(values, row, 0, colNames, "DNB CVB Rating").Trim().ToUpper();
                        }
                    }

                    if (!IMW_File_Data.ContainsKey(contractData.m_sSecurityID_LL))
                    {
                        IMW_File_Data.Add(contractData.m_sSecurityID_LL, new Instrument_Bond_OriginalData_List());
                        numberOfSwaps++;
                    }
                    IMW_File_Data[contractData.m_sSecurityID_LL].Add(contractData);
                    if (!contractData.m_bLookthroughData)
                    {// IMW cash flow are only for DIMENSION DATA: 
                        if (!SecurityTypesList[securityType_LL].Contains(contractData.m_sLeg))
                        {
                            SecurityTypesList[securityType_LL].Add(contractData.m_sLeg);
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                errors.AddError("Fout tijdens inlezen BOND positie in rij " + row + " : " + exc.Message);
                return new PositionList();
            }
            // Read Cash Flow file:
            // risk neutral
            Dictionary<string, CFixedIncomeCashFlowData> FixedIncomeCashFlowData_RiskNeutral = ReadCashFlowFile_IMW(dtNow, CashFlow_FileName,
                SecurityTypesList, CashFlowType.RiskNeutral, errors);
            // rente typisch
            Dictionary<string, CFixedIncomeCashFlowData> FixedIncomeCashFlowData_RiskRente = ReadCashFlowFile_IMW(dtNow, CashFlow_FileName,
                SecurityTypesList, CashFlowType.RiskRente, errors);

            // Create the position list:
            PositionList positions = new PositionList();
            try
            {
                foreach (KeyValuePair<string, Instrument_Bond_OriginalData_List> CSecurityData in IMW_File_Data)
                {
                    string m_sSecurityID_LL = CSecurityData.Key;
                    Instrument_Bond_OriginalData_List DataList = CSecurityData.Value;
                    foreach (Instrument_Bond_OriginalData CPositionData in DataList)
                    {
                        string UniqSecurityID = CPositionData.m_sUniquePositionId;
                        CFixedIncomeCashFlowData[] FixedIncomeObj;
                        if (CPositionData.m_bLookthroughData)
                        {
                            FixedIncomeObj = null;
                        }
                        else
                        {
                            if (!FixedIncomeCashFlowData_RiskRente.ContainsKey(UniqSecurityID))
                            {
                                string message = "No (Rente Typisch) Cash Flow for Row (" + CPositionData.m_dRow
                                    + ") Portfolio (" + CPositionData.m_sPortfolioID
                                    + ") Security (" + CPositionData.m_sSecurityID_LL
                                    + ") Leg (" + CPositionData.m_sLeg + ")"
                                    + ") Volume (" + CPositionData.m_fNominal + ")";
                                if (true ||
                                    MessageBox.Show(message + ". Bestand alsnog verwerken?", "Ongeldige rapportagedatum", MessageBoxButtons.OKCancel) == DialogResult.OK)
                                {
                                    errors.AddWarning(message + " Gebruiker heeft melding genegeerd");
                                    continue;
                                }
                                else
                                {
                                    errors.AddError(message + " Gebruiker heeft verwerking gestopt");
                                    return new PositionList();
                                }
                            }
                            if (!FixedIncomeCashFlowData_RiskNeutral.ContainsKey(UniqSecurityID))
                            {
                                string message = "No (Risk Neutral) Cash Flow for Row (" + CPositionData.m_dRow
                                    + ") Portfolio (" + CPositionData.m_sPortfolioID
                                    + ") Security (" + CPositionData.m_sSecurityID_LL
                                    + ") Leg (" + CPositionData.m_sLeg + ")"
                                    + ") Volume (" + CPositionData.m_fNominal + ")";
                                errors.AddError(message + " Gebruiker heeft verwerking gestopt");
                                return new PositionList();
                            }
                            FixedIncomeObj = new CFixedIncomeCashFlowData[2];
                            FixedIncomeObj[(int)CashFlowType.RiskRente] = FixedIncomeCashFlowData_RiskRente[UniqSecurityID];
                            FixedIncomeObj[(int)CashFlowType.RiskNeutral] = FixedIncomeCashFlowData_RiskNeutral[UniqSecurityID];
                        }
                        Position position = new Position();
                        position.m_sDATA_Source = "IMW";
                        position.m_sRow = CPositionData.m_dRow.ToString();
                        position.m_bIsLookThroughPosition = CPositionData.m_bLookthroughData;

                        if ("FUNDING" != CPositionData.m_sPortfolioPurpose)
                        {
                            position.m_sBalanceType = "Assets";
                            position.m_sGroup = "Bonds";
                        }
                        else
                        {
                            position.m_sBalanceType = "Liabilities";
                            position.m_sGroup = "Bonds";
                        }
                        position.m_sUniquePositionId = CPositionData.m_sUniquePositionId;
                        position.m_sSelectieIndex_LL = CPositionData.m_sSelectieIndex_LL;
                        position.m_sSecurityType_LL = CPositionData.m_sSecurityType_LL;
                        position.m_bICO = CPositionData.m_bICO;
                        position.m_sScope3 = CPositionData.m_sScope3;
                        position.m_sSecurityID = CPositionData.m_sSecurityID;
                        position.m_sSecurityName = CPositionData.m_sSecurityName;
                        position.m_sSecurityID_LL = CPositionData.m_sSecurityID_LL;
                        position.m_sSecurityName_LL = CPositionData.m_sSecurityName_LL;
                        position.m_sPositionId = position.m_sSecurityID_LL;
                        position.m_sPortfolioId = CPositionData.m_sPortfolioID;

                        position.m_sLegId = CPositionData.m_sLeg;
                        position.m_sCurrency = CPositionData.m_sCurrency;
                        position.m_sCountryCurrency = CPositionData.m_sCurrencyCountry;
                        position.m_fFxRate = CPositionData.m_fFxRate;
                        position.m_fVolume = CPositionData.m_fNominal;
                        position.m_sRiskClass = CPositionData.m_sNACEcode;
                        position.m_sCIC = CPositionData.m_sCIC_ID;
                        position.m_sCIC_LL = CPositionData.m_sCIC_ID_LL;
                        position.m_sCIC_SCR = position.m_sCIC_LL;
                        position.m_bGovGuarantee = CPositionData.m_bGovGuarantee;
                        position.m_bEEA = CPositionData.m_bEEA;
                        position.m_fCollateralCoveragePercentage = CPositionData.m_fCollateral;
                        position.m_dSecuritisationType = CPositionData.m_dSecuritisationType;
                        position.m_sSecurityCreditQuality = CPositionData.m_sSecurityCreditQuality;
                        position.m_dSecurityCreditQuality = CPositionData.m_dSecurityCreditQuality;
                        position.m_sIssuerCreditQuality = CPositionData.m_sSecurityCreditQuality;
                        position.m_dIssuerCreditQuality = CPositionData.m_dIssuerCreditQuality;
                        position.m_fModifiedDuration = CPositionData.m_fModifiedDurationCorrected;
                        position.m_fSpreadDuration = CPositionData.m_fSpreadDurationCorrected;
                        DateTime Maturity = CPositionData.m_dtReport.Value;
                        if (CPositionData.m_fExpiryDate > 0)
                        {
                            Maturity = getExcelDate_From_DoubleDate(CPositionData.m_fExpiryDate).Value;
                        }
                        else
                        {
                            Maturity = Maturity.AddYears(3);
                        }
                        Instrument_Bond instrument = new Instrument_Bond(Maturity, CPositionData, FixedIncomeObj);
                        if (CPositionData.m_bCallable)
                        {
                            instrument.m_bCallable = true;
                            instrument.m_CallDate = getExcelDate_From_DoubleDate(CPositionData.m_fCallDate).Value;
                        }

                        string ccyTranslated = Position.TranslateCurrency_Curve(scenarioZeroCurves, position.m_sCurrency);
                        zeroCurve = scenarioZeroCurves[ccyTranslated];
                        try
                        {
                            IndexCPI CPIindex = null;
                            if (BondType.INFLATION == CPositionData.m_BondType)
                            {
                                if (scenarioIndexCPI_List.Count > 0)
                                {
                                    CPIindex = scenarioIndexCPI_List[ccyTranslated];
                                }
                            }
                            instrument.Init(dtNow, zeroCurve, CPIindex, CPositionData.m_fMarketValue); // Spread floored at -10%
                            position.m_Instrument = instrument;
                        }
                        catch (Exception exc)
                        {
                            string errorMassage = exc.Message + "; the run has stopped: Initialization of the bond (" + CPositionData.m_sUniquePositionId + ") has failed!";
                            errors.AddWarning(errorMassage);
                            throw new IOException(errorMassage);
                        }

                        position.m_sSecurityID = CPositionData.m_sSecurityID;
                        position.m_sCIC = CPositionData.m_sCIC_ID;
                        position.m_fAccruedInterest_LL = CPositionData.m_fAccruedInterest_LL;
                        position.m_fAccruedDividend_LL = 0;
                        position.m_sAccount = CPositionData.m_sAccount;
                        position.m_sAccount_LL = CPositionData.m_sAccount_LL;
                        position.m_sECAP_Category_LL = CPositionData.m_sECAP_Category_LL;
                        if (position.m_sSelectieIndex_LL.Substring(10, 1) == "1") // 11th digit for Spread risk
                        {// not used  for this product
                            position.m_SpreadRiskData = new PositionSpreadRiskData();
                            position.m_SpreadRiskData.m_sSelectieIndex_LL = position.m_sSelectieIndex_LL;
                            position.m_SpreadRiskData.m_sCIC_LL = position.m_sCIC_LL;
                            position.m_SpreadRiskData.m_dSecuritisationType = position.m_dSecuritisationType;
                            position.m_SpreadRiskData.m_sSecurityName_LL = position.m_sSecurityName_LL;
                            position.m_SpreadRiskData.m_sInstrumentID_LL = position.m_sSecurityID_LL;
                            position.m_SpreadRiskData.m_sPortfolioID = position.m_sPortfolioId;
                            position.m_SpreadRiskData.m_sCountryCurrency = position.m_sCountryCurrency;
                            position.m_SpreadRiskData.m_sCurrency = position.m_sCurrency;
                            position.m_SpreadRiskData.m_bEEA = position.m_bEEA;
                            position.m_SpreadRiskData.m_dSecurityCreditQuality = position.m_dSecurityCreditQuality;
                            position.m_SpreadRiskData.m_bGovGuarantee = position.m_bGovGuarantee;
                            position.m_SpreadRiskData.m_fCollateral = position.m_fCollateralCoveragePercentage; // in %
                            position.m_SpreadRiskData.m_fModifiedDuration = position.m_fModifiedDuration;
                            position.m_SpreadRiskData.m_fSpreadDuration = position.m_fSpreadDuration;
                            SCRSpreadModel.WhichProduct(position.m_SpreadRiskData);
                        }

                        if (CPositionData.m_dtReport > instrument.m_MaturityDate)
                        {
                            position.m_bHasMessage = true;
                            position.m_sMessage = " Warning : Instrument is matured. Fair Value wordt 0 voor alle scenarios.";
                            errors.Add("Row = " + position.m_sRow + " ID = " + position.m_sPositionId + position.m_sMessage);
                        }
                        else
                        {
                            if (!CPositionData.m_bLookthroughData
                                && (instrument.m_fMarketPrice != 0)
                                && (instrument.m_OriginalCashFlowData[(int)CashFlowType.RiskRente].m_CashFlowSched.Count() == 0))
                            {
                                position.m_bHasMessage = true;
                                position.m_sMessage = " Warning : CleanValue is niet nul, maar er zijn geen kasstromen. Fair Value wordt 0 voor alle scenarios.";
                                errors.Add("Row = " + position.m_sRow + " ID = " + position.m_sPositionId + position.m_sMessage);
                            }
                            else if (!instrument.m_bImpliedSpreadFound)
                            {
                                position.m_bHasMessage = true;
                                position.m_sMessage = " Warning : Spread kan niet worden bepaald. Fair Value wordt gebruikt voor alle scenarios.";
                                errors.Add("Row = " + position.m_sRow + " ID = " + position.m_sPositionId + position.m_sMessage);
                            }
                        }
                        // Set the static data of current and previous period:
                        CBondDebugData p = model.getDebugData_Bond(sPeriod, position);
                        instrument.m_BondDebugData_CurrentPeriod = p;
                        instrument.m_BondDebugData_PrevPeriod = model.getLinkedPosition_Bond(p);
                        positions.AddPosition(position);
                    }
                }
            }
            catch (Exception exc)
            {
                errors.AddError("Error during the initialization of the bond position in row " + row + " : " + exc.Message);
                positions = new PositionList();
            }

            return positions;
        }
        public PositionList ReadSwaptionPositions_IMW(DateTime dtNow, string fileName,
            ScenarioList scenarios, string currency, Boolean hullWhiteModel, ErrorList errors)
        {
            const int RowStart = 2;

            TotalRisk.Utilities.Scenario baseScenario = scenarios.getScenarioFairValue();
            Curve zeroCurve = baseScenario.m_YieldCurves.ByName(currency).m_Curve;
            double hullWhiteA = baseScenario.GetHullWhiteMeanReversion(currency);
            double hullWhiteSigma = baseScenario.GetHullWhiteVolatility(currency);
            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, "Integrale aanlevering IMW maand", "A1");
            Dictionary<string, int> colNames = HeaderNamesColumns(values);
            PositionList positions = new PositionList(values.GetUpperBound(DimensionRow));
            int year, month, day, row = 0;
            try
            {
                for (row = RowStart; row <= values.GetUpperBound(DimensionRow); row++)
                {

                    string reportCode = ReadFieldAsString(values, row, 0, colNames, "SelectieIndex LL");
                    if (reportCode.Substring(2, 1) != "1") // third digit
                    {
                        continue;
                    }
                    string cic_ID = ReadFieldAsString(values, row, 0, colNames, "Cic Id Ll").ToUpper();
                    string cic_ID_last2 = cic_ID.Substring(2, 2);
                    if (cic_ID_last2 == "B6" || cic_ID_last2 == "C6")
                    {
                    }
                    else
                    {
                        continue;
                    }

                    Instrument_Swaption_OriginalData swaptionData = new Instrument_Swaption_OriginalData();
                    swaptionData.m_sSelectieIndex_LL = reportCode;
                    swaptionData.m_sScope3 = ReadFieldAsString(values, row, 0, colNames, "Ecs Cons Ecap asr");
                    swaptionData.m_sPortfolioID = ReadFieldAsString(values, row, 0, colNames, "Portfolio Id");
                    swaptionData.m_sCICLL = cic_ID;
                    if (swaptionData.m_sPortfolioID == "")
                    {
                        continue;
                    }
                    swaptionData.m_fMarketValue = ReadFieldAsDouble(values, row, 0, colNames, "Market Value Eur Ll");
                    swaptionData.m_fFxRate = 1.0 / ReadFieldAsDouble(values, row, 0, colNames, "Fx Rate Qc Pc Laagste Lt Level");
                    swaptionData.m_fNominal = ReadFieldAsDouble(values, row, 0, colNames, "BalNomVal LT");
                    swaptionData.m_sType = ReadFieldAsString(values, row, 0, colNames, "Derivaten Type Ll");
                    swaptionData.m_fStrike = ReadFieldAsDouble(values, row, 0, colNames, "Strike Price Laagste Level") / 100;

                    swaptionData.m_fSwaptionExpiry = ReadFieldAsDouble(values, row, 0, colNames, "Maturity Call LL");
                    swaptionData.m_fSwapExpiry = ReadFieldAsDouble(values, row, 0, colNames, "Maturity LL");
                    swaptionData.m_sInstrumentType = ReadFieldAsString(values, row, 0, colNames, "Instrument Type Ll").ToUpper();
                    swaptionData.m_sSecurityID = ReadFieldAsString(values, row, 0, colNames, "Security Id Ll");
                    swaptionData.m_sSecurityName = ReadFieldAsString(values, row, 0, colNames, "Security Name Ll");
                    string securityType = ReadFieldAsString(values, row, 0, colNames, "Security Type Ll").ToUpper();
                    SwaptionType type = cic_ID_last2.StartsWith("C") ? SwaptionType.Receiver : SwaptionType.Payer;
                    DateTime? dtExpiry = getExcelDate_From_DoubleDate(swaptionData.m_fSwaptionExpiry);
                    DateTime? dtMaturity = getExcelDate_From_DoubleDate(swaptionData.m_fSwapExpiry);
                    swaptionData.m_fSwaptionVolatility = ReadFieldAsDouble(values, row, 0, colNames, "Volatility dim") / 100;

                    // credit quality steps Group Vounterparty:
                    swaptionData.m_sGroupCounterpartyName = ReadFieldAsString(values, row, 0, colNames, "Groep tegenpartij naam Ll");
                    swaptionData.m_sGroupCounterpartyLEI = ReadFieldAsString(values, row, 0, colNames, "Groep tegenpartij LEI Ll");
                    string test = ReadFieldAsString(values, row, 0, colNames, "Groep tegenpartij Credit Quality Step Ll");
                    swaptionData.m_sGroupCounterpartyCQS = test;
                    if (test == "" || test.Substring(0, 1) == "NR")
                    {//  seven as NR
                        swaptionData.m_dGroupCounterpartyCQS = 7;
                    }
                    else
                    {
                        swaptionData.m_dGroupCounterpartyCQS = Convert.ToInt32(test.Substring(0, 1));
                    }

                    bool cashSettled = true; // FM approved temporaly solution 2015-1-8
                                             //                    bool cashSettled = false; // FM approved temporaly solution 2020-02-26
                    Instrument_Swaption instrument = new Instrument_Swaption(swaptionData, type, cashSettled, swaptionData.m_fStrike, dtExpiry.Value, dtMaturity.Value);
                    double marketValueNormalized = 0;
                    if (swaptionData.m_fNominal != 0)
                    {
                        marketValueNormalized = swaptionData.m_fMarketValue / swaptionData.m_fNominal;
                    }
                    if (hullWhiteModel)
                    {
                        instrument.Init(dtNow, zeroCurve, marketValueNormalized, hullWhiteA, hullWhiteSigma);
                    }
                    else
                    {
                        instrument.Init(dtNow, zeroCurve, marketValueNormalized, swaptionData.m_fSwaptionVolatility);
                    }

                    Position position = new Position();
                    position.m_sSelectieIndex_LL = reportCode;
                    position.m_sDATA_Source = "IMW";
                    position.m_sRow = row.ToString();
                    position.m_bIsLookThroughPosition = false;

                    position.m_Instrument = instrument;
                    position.m_sBalanceType = "Assets";
                    position.m_sGroup = "Swaptions";
                    position.m_sSecurityType_LL = swaptionData.m_sInstrumentType;
                    position.m_sCIC_LL = swaptionData.m_sCICLL;
                    position.m_sCIC_SCR = position.m_sCIC_LL;
                    position.m_sScope3 = swaptionData.m_sScope3;
                    position.m_sScope3 = ScopeData.getScopeFormated(position.m_sScope3);
                    position.m_sSecurityID_LL = swaptionData.m_sSecurityID;
                    position.m_sSecurityName_LL = swaptionData.m_sSecurityName;
                    position.m_sPortfolioId = swaptionData.m_sPortfolioID;
                    position.m_sPositionId = position.m_sSecurityID_LL;
                    position.m_fFxRate = swaptionData.m_fFxRate;
                    position.m_fVolume = swaptionData.m_fNominal;
                    position.m_fSCR_weight = Math.Max(0, Math.Min(1, instrument.getMaturity(dtNow)));

                    position.m_sSecurityID = ReadFieldAsString(values, row, 0, colNames, "Security Id");
                    position.m_sCIC = ReadFieldAsString(values, row, 0, colNames, "Cic Id").ToUpper();
                    position.m_fFairValue = swaptionData.m_fMarketValue;
                    position.m_fAccruedInterest_LL = ReadFieldAsDouble(values, row, 0, colNames, "Accrued Interest Ll");
                    position.m_fAccruedDividend_LL = ReadFieldAsDouble(values, row, 0, colNames, "Accrued Dividend Ll");
                    position.m_sAccount = ReadFieldAsString(values, row, 0, colNames, "RDS-STA account");
                    position.m_sAccount_LL = ReadFieldAsString(values, row, 0, colNames, "RDS-STA account LT");
                    position.m_sECAP_Category_LL = ReadFieldAsString(values, row, 0, colNames, "ECAP Category Ll");


                    positions.AddPosition(position);

                    System.Windows.Forms.Application.DoEvents();
                }
            }
            catch (Exception exc)
            {
                errors.AddError("Fout tijdens inlezen swaption positie in rij " + row + " : " + exc.Message);
            }

            return positions;
        }
        public PositionList ReadSwapPositions_IMW(DateTime dtNow, string IMW_FileName, string CashFlow_FileName, ScenarioList scenarios, ErrorList errors)
        {
            Dictionary<string, List<string>> SecurityTypesList = new Dictionary<string, List<string>>();
            SecurityTypesList.Add("ZCISW", new List<string>());
            SecurityTypesList.Add("IRS", new List<string>());
            const int RowStart = 2;
            // Defne the base curves:
            TotalRisk.Utilities.Scenario baseScenario = scenarios.getScenarioFairValue();
            CurveList scenarioZeroCurves = new CurveList();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
            {
                string ccy = scenarioCurve.m_sName;
                scenarioZeroCurves.Add(ccy.ToUpper(), scenarioCurve.m_Curve);
            }
            CurveList scenarioEONIA_SpreadCurves = new CurveList();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_EONIA_SpreadCurves)
            {
                string ccy = scenarioCurve.m_sName;
                scenarioEONIA_SpreadCurves.Add(ccy.ToUpper(), scenarioCurve.m_Curve);
            }
            CurveList scenarioZeroInflationCurves = new CurveList();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_InflationCurves)
            {
                string ccy = scenarioCurve.m_sName;
                scenarioZeroInflationCurves.Add(ccy.ToUpper(), scenarioCurve.m_Curve);
            }
            SortedList<string, IndexCPI> scenarioIndexCPI_List = new SortedList<string, IndexCPI>();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_InflationCurves)
            {
                string ccy = scenarioCurve.m_sName;
                scenarioZeroInflationCurves.Add(ccy.ToUpper(), scenarioCurve.m_Curve);
                IndexCPI CPIobject = new IndexCPI();
                CPIobject.SetInflationInstance(dtNow, 100, scenarioCurve.m_Curve, 100);
                scenarioIndexCPI_List.Add(ccy.ToUpper(), CPIobject);
            }

            Curve[] zeroCurve = new Curve[2];
            Curve[] EONIASpreadCurve = new Curve[2];
            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(IMW_FileName, "Integrale aanlevering IMW maand", "A1");
            Dictionary<string, int> colNames = HeaderNamesColumns(values);
            // Scope : Security Id Ll : (leg, leg data)
            Dictionary<string, Dictionary<string, Instrument_Swap_OriginalData_List>> OriginalData =
                new Dictionary<string, Dictionary<string, Instrument_Swap_OriginalData_List>>();
            int row = 0;
            int numberOfSwaps = 0;
            try
            {
                for (row = RowStart; row <= values.GetUpperBound(DimensionRow); row++)
                {

                    string reportCode = ReadFieldAsString(values, row, 0, colNames, "SelectieIndex LL");
                    string securityType_LL = ReadFieldAsString(values, row, 0, colNames, "Security Type Ll").ToUpper().Trim();
                    string cic_ID = ReadFieldAsString(values, row, 0, colNames, "Cic Id Ll").ToUpper();
                    string cic_ID_last2 = cic_ID.Substring(2, 2);
                    if (reportCode.Substring(3, 1) == "1") // forth digit
                    {
                    }
                    else if (reportCode.Substring(0, 1) == "1" && // first digit
                        (cic_ID_last2 == "E2" || cic_ID_last2 == "D9"))
                    {
                    }
                    else
                    {
                        continue;
                    }
                    SwapType swapType;
                    if (cic_ID_last2 == "D1")
                    {
                        swapType = SwapType.IRS;
                    }
                    else if (cic_ID_last2 == "D3")
                    {
                        swapType = SwapType.CurrencySwap;
                    }
                    else if (cic_ID_last2 == "E2")
                    {
                        swapType = SwapType.FX_FORWARD;
                    }
                    else if (cic_ID_last2 == "D9" && "ZCISW" == securityType_LL)
                    {
                        swapType = SwapType.ZC_InflationSwap;
                    }
                    else
                    {
                        continue;
                    }

                    Instrument_Swap_OriginalData swapData = new Instrument_Swap_OriginalData();
                    double dateValue = ReadFieldAsDouble(values, row, 0, colNames, "Reporting Date");
                    swapData.m_dtReport = getExcelDate_From_DoubleDate(dateValue);
                    swapData.m_dRow = row;
                    swapData.m_sSwapType = swapType;
                    swapData.m_sSelectieIndex_LL = reportCode;
                    swapData.m_sSecurityType = securityType_LL;
                    swapData.m_sSecurityID_LL = ReadFieldAsString(values, row, 0, colNames, "Security Id Ll");
                    swapData.m_sSecurityName = ReadFieldAsString(values, row, 0, colNames, "Security Name Ll");
                    if ("ZC10010" == swapData.m_sSecurityID_LL)
                    {
                        row += 0;
                    }

                    swapData.m_sCIC_ID_LL = cic_ID;
                    swapData.m_sScope3 = "";
                    swapData.m_sPortfolioID = ReadFieldAsString(values, row, 0, colNames, "Portfolio Id");
                    swapData.m_sScope3 = ReadFieldAsString(values, row, 0, colNames, "Ecs Cons Ecap asr");
                    if (swapData.m_sScope3 == "")
                    {
                        swapData.m_sScope3 = "0";
                    }
                    else
                    {
                        if (swapData.m_sScope3.Contains("_"))
                        {
                            swapData.m_sScope3 = swapData.m_sScope3.Substring(0, 4);
                        }
                    }
                    // Start Date:
                    swapData.m_fStartDate = ReadFieldAsDouble(values, row, 0, colNames, "Initial Start Date");
                    swapData.m_StartDate = getExcelDate_From_DoubleDate(swapData.m_fStartDate);

                    if (null == swapData.m_StartDate && (SwapType.IRS == swapType || SwapType.CurrencySwap == swapType))
                    {
                        throw new Exception("Swap/FX Forward ( " + swapData.m_sSecurityID_LL + " ) has no data in column (Initial Start Date), row =  " + row);
                    }
                    if (swapData.m_StartDate > swapData.m_dtReport)
                    {
                        swapData.m_bForwardStartSwap = true;
                    }
                    else
                    {
                        swapData.m_bForwardStartSwap = false;
                    }
                    // Maturity:
                    swapData.m_fExpiryDate = ReadFieldAsDouble(values, row, 0, colNames, "Maturity LL");
                    swapData.m_ExpiryDate = getExcelDate_From_DoubleDate(swapData.m_fExpiryDate);
                    if (null == swapData.m_ExpiryDate)
                    {
                        throw new Exception("Swap/FX Forward ( " + swapData.m_sSecurityID_LL + " ) has no data in column (Maturity LL), row =  " + row);
                    }
                    // Nominal:
                    string test = ReadFieldAsString(values, row, 0, colNames, "Balnomval qc");
                    if (test.Equals(""))
                    {
                        throw new Exception("Swap/FX ( " + swapData.m_sSecurityID_LL + " ) has no data in column (Balnomval qc), row =  " + row);
                    }
                    swapData.m_fNominal = ReadFieldAsDouble(values, row, 0, colNames, "Balnomval qc");
                    // Market Value:
                    //                    swapData.m_fMarketValue = ReadFieldAsDouble(values, row, 0, colNames, "Market Value Eur Ll"); // not availible for FX-Forward
                    test = ReadFieldAsString(values, row, 0, colNames, "Ccy Exposure Pc Laagste Level");
                    if (test.Equals(""))
                    {
                        throw new Exception("Swap/FX ( " + swapData.m_sSecurityID_LL + " ) has no data in column (Ccy Exposure Pc Laagste Level), row =  " + row);
                    }
                    swapData.m_fFxRate = 1.0 / ReadFieldAsDouble(values, row, 0, colNames, "Fx Rate Qc Pc Laagste Lt Level");
                    swapData.m_fMarketValue = ReadFieldAsDouble(values, row, 0, colNames, "Ccy Exposure Pc Laagste Level"); // in EUR in 2021M02 included (as of 2021Q1 it is in Product currency)
                    swapData.m_fMarketValue *= swapData.m_fFxRate; // as of 20221Q1
                    swapData.m_fAccruedInterest_LL = ReadFieldAsDouble(values, row, 0, colNames, "Accrued Interest Ll");
                    swapData.m_sType = ReadFieldAsString(values, row, 0, colNames, "Derivaten Type Ll");
                    swapData.m_PaymentType = swapData.m_fNominal > 0 ? SwapPaymentType.Receiver : SwapPaymentType.Payer;

                    swapData.m_dLegID = ReadFieldAsInt(values, row, 0, colNames, "Leg no");
                    swapData.m_sCouponType = ReadFieldAsString(values, row, 0, colNames, "Coupon Type Laagste Lt Level").ToUpper();
                    if (SwapType.FX_FORWARD == swapType)
                    {
                        swapData.m_bFixedLeg = true;
                    }
                    else if (SwapType.ZC_InflationSwap == swapType)
                    {
                        if (1 == swapData.m_dLegID)
                        {
                            swapData.m_bFixedLeg = false;
                        }
                        else
                        {
                            swapData.m_bFixedLeg = true;
                        }
                    }
                    else if ("FIXED" == swapData.m_sCouponType)
                    {
                        swapData.m_bFixedLeg = true;
                    }
                    else
                    {
                        swapData.m_bFixedLeg = false;
                    }

                    swapData.m_fFixedLegRate = ReadFieldAsDouble(values, row, 0, colNames, "Coupon Perc Laagste Lt Level") / 100;
                    swapData.m_fFloatingLegRate = swapData.m_fFixedLegRate;

                    swapData.m_sCurrency = ReadFieldAsString(values, row, 0, colNames, "Currency Laagste Lt Level").ToUpper().Trim();

                    swapData.m_dFrequency = ReadFieldAsInt(values, row, 0, colNames, "Coupon Frequency Laagste Level");

                    swapData.m_sSecurityID = ReadFieldAsString(values, row, 0, colNames, "Security Id");
                    swapData.m_sCIC_ID = ReadFieldAsString(values, row, 0, colNames, "Cic Id").ToUpper();
                    swapData.m_sAccount = ReadFieldAsString(values, row, 0, colNames, "RDS-STA account");
                    swapData.m_sAccount_LL = ReadFieldAsString(values, row, 0, colNames, "RDS-STA account LT");
                    swapData.m_sECAP_Category_LL = ReadFieldAsString(values, row, 0, colNames, "ECAP Category Ll");
                    // credit quality steps Group Vounterparty:
                    swapData.m_sGroupCounterpartyName = ReadFieldAsString(values, row, 0, colNames, "Groep tegenpartij naam Ll");
                    swapData.m_sGroupCounterpartyLEI = ReadFieldAsString(values, row, 0, colNames, "Groep tegenpartij LEI Ll");
                    test = ReadFieldAsString(values, row, 0, colNames, "Groep tegenpartij Credit Quality Step Ll");
                    swapData.m_sGroupCounterpartyCQS = test;
                    if (test == "" || test.Substring(0, 1) == "NR")
                    {//  seven as NR
                        swapData.m_dGroupCounterpartyCQS = 7;
                    }
                    else
                    {
                        swapData.m_dGroupCounterpartyCQS = Convert.ToInt32(test.Substring(0, 1));
                    }


                    if (!OriginalData.ContainsKey(swapData.m_sScope3))
                    {
                        OriginalData.Add(swapData.m_sScope3, new Dictionary<string, Instrument_Swap_OriginalData_List>());
                    }
                    Dictionary<string, Instrument_Swap_OriginalData_List> pDict = OriginalData[swapData.m_sScope3];
                    if (!pDict.ContainsKey(swapData.m_sSecurityID_LL))
                    {
                        pDict.Add(swapData.m_sSecurityID_LL, new Instrument_Swap_OriginalData_List());
                        numberOfSwaps++;
                    }
                    swapData.m_sUniquePositionId = swapData.m_sPortfolioID + "_" + swapData.m_sSecurityID_LL + "_" + swapData.m_dLegID;
                    pDict[swapData.m_sSecurityID_LL].Add(swapData);
                    if (SwapType.ZC_InflationSwap == swapType)
                    {
                        if (!SecurityTypesList["ZCISW"].Contains(swapData.m_dLegID.ToString()))
                        {
                            SecurityTypesList["ZCISW"].Add(swapData.m_dLegID.ToString());
                        }
                        if (!SecurityTypesList["IRS"].Contains(swapData.m_dLegID.ToString()))
                        {
                            SecurityTypesList["IRS"].Add(swapData.m_dLegID.ToString());
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                errors.AddError("Fout tijdens inlezen swap positie in rij " + row + " : " + exc.Message);
                return new PositionList();
            }
            // Read Cash Flow file:
            // risk neutral
            Dictionary<string, CFixedIncomeCashFlowData> FixedIncomeCashFlowData_RiskNeutral = ReadCashFlowFile_IMW(dtNow, CashFlow_FileName,
                SecurityTypesList, CashFlowType.RiskNeutral, errors);
            // rente typisch
            Dictionary<string, CFixedIncomeCashFlowData> FixedIncomeCashFlowData_RiskRente = ReadCashFlowFile_IMW(dtNow, CashFlow_FileName,
                SecurityTypesList, CashFlowType.RiskRente, errors);

            // Create the position list:
            PositionList positions = new PositionList(numberOfSwaps);
            int n = 0;
            int IMW_Row1 = 0;
            int IMW_Row2 = 0;
            try
            {
                foreach (Dictionary<string, Instrument_Swap_OriginalData_List> pScopeData in OriginalData.Values)
                {
                    foreach (Instrument_Swap_OriginalData_List pSwapLegsData in pScopeData.Values)
                    {
                        n++;
                        IMW_Row1 = 0;
                        IMW_Row2 = 0;
                        if (34 == n)
                        {
                            row += 0;
                        }
                        Instrument_Swap_OriginalData[] pArray = pSwapLegsData.ToArray();
                        if (pArray.Length != 2)
                        {
                            throw new Exception("Swap ( " + pArray[0].m_sSecurityID_LL + " ) has only " + pArray.Length + "legs, but must have 2 legs! ");
                        }
                        IMW_Row1 = pArray[0].m_dRow;
                        IMW_Row2 = pArray[1].m_dRow;
                        if (pArray[0].m_sCIC_ID_LL != pArray[1].m_sCIC_ID_LL)
                        {
                            throw new Exception("Swap ( " + pArray[0].m_sSecurityID_LL + " ) has the legs of different CIC ID: leg 1 =  " + pArray[0].m_sCIC_ID_LL + " leg 2 = " + pArray[1].m_sCIC_ID_LL);
                        }
                        string cic_ID_last2 = pArray[0].m_sCIC_ID_LL.Substring(2, 2);
                        //                        if ("IRSW815597" == pArray[0].m_sSecurityID_LL)
                        //                        {
                        //                            n += 0;
                        //                        }

                        SwapType swapType;
                        if (SwapType.IRS == pArray[0].m_sSwapType)
                        {
                            if (pArray[0].m_bFixedLeg == pArray[1].m_bFixedLeg)
                            {
                                throw new Exception(" IR Swap ( " + pArray[0].m_sSecurityID_LL + " ) has both legs of the same type: " + ((pArray[0].m_bFixedLeg) ? " FIXED!" : " FLOAT!"));
                            }
                            if (pArray[0].m_sCurrency != pArray[1].m_sCurrency)
                            {
                                throw new Exception(" IR Swap ( " + pArray[0].m_sSecurityID_LL + " ) has legs of different currencies: leg_1 (" + pArray[0].m_sCurrency + ") leg_2 (" + pArray[0].m_sCurrency + ")");
                            }
                            swapType = pArray[0].m_sSwapType;
                        }
                        else if (SwapType.CurrencySwap == pArray[0].m_sSwapType)
                        {
                            if (!pArray[0].m_bFixedLeg || !pArray[1].m_bFixedLeg)
                            {
                                throw new Exception(" Currency Swap ( " + pArray[0].m_sSecurityID_LL + " ) has one or both legs being FLOAT, which is not implemented here!");
                            }
                            if (pArray[0].m_sCurrency == pArray[1].m_sCurrency)
                            {
                                throw new Exception(" Currency Swap ( " + pArray[0].m_sSecurityID_LL + " ) has both legs of the same currency: (" + pArray[0].m_sCurrency + ")");
                            }
                            swapType = pArray[0].m_sSwapType;
                        }
                        else if (SwapType.FX_FORWARD == pArray[0].m_sSwapType)
                        {
                            if (!pArray[0].m_bFixedLeg || !pArray[1].m_bFixedLeg)
                            {
                                throw new Exception(" FX_FORWARD ( " + pArray[0].m_sSecurityID_LL + " ) has one or both legs being FLOAT, which is not implemented here!");
                            }
                            if (pArray[0].m_sCurrency == pArray[1].m_sCurrency)
                            {
                                throw new Exception(" FX_FORWARD ( " + pArray[0].m_sSecurityID_LL + " ) has both legs of the same currency: (" + pArray[0].m_sCurrency + ")");
                            }
                            swapType = pArray[0].m_sSwapType;
                        }
                        else if (SwapType.ZC_InflationSwap == pArray[0].m_sSwapType)
                        {
                            if (pArray[0].m_bFixedLeg == pArray[1].m_bFixedLeg)
                            {
                                throw new Exception(" ZC Inflation Swap ( " + pArray[0].m_sSecurityID_LL + " ) has both legs of the same type: " + ((pArray[0].m_bFixedLeg) ? " FIXED!" : " FLOAT!"));
                            }
                            if (pArray[0].m_sCurrency != pArray[1].m_sCurrency)
                            {
                                throw new Exception(" ZC Inflation Swap ( " + pArray[0].m_sSecurityID_LL + " ) has legs of different currencies: leg_1 (" + pArray[0].m_sCurrency + ") leg_2 (" + pArray[0].m_sCurrency + ")");
                            }
                            swapType = pArray[0].m_sSwapType;
                        }
                        else
                        {
                            continue;
                        }
                        Instrument_Swap_OriginalData swapDataFixedLeg, swapDataFloatLeg;
                        // IRS : normal interst rate swap
                        if (swapType == SwapType.IRS)
                        {
                            if (pArray[0].m_bFixedLeg)
                            {
                                swapDataFixedLeg = pArray[0];
                                swapDataFloatLeg = pArray[1];
                            }
                            else
                            {
                                swapDataFixedLeg = pArray[1];
                                swapDataFloatLeg = pArray[0];
                            }
                        }
                        else if (swapType == SwapType.ZC_InflationSwap)
                        {
                            if (pArray[0].m_bFixedLeg)
                            {
                                swapDataFixedLeg = pArray[0];
                                swapDataFloatLeg = pArray[1];
                            }
                            else
                            {
                                swapDataFixedLeg = pArray[1];
                                swapDataFloatLeg = pArray[0];
                            }
                        }
                        else // Currencu swap and FX Forward
                        {
                            if (pArray[0].m_sCurrency == "EUR")
                            {
                                swapDataFixedLeg = pArray[0];
                                swapDataFloatLeg = pArray[1];
                            }
                            else if (pArray[1].m_sCurrency == "EUR")
                            {
                                swapDataFixedLeg = pArray[1];
                                swapDataFloatLeg = pArray[0];
                            }
                            else if (pArray[0].m_dLegID < pArray[1].m_dLegID)
                            {
                                swapDataFixedLeg = pArray[0];
                                swapDataFloatLeg = pArray[1];
                            }
                            else
                            {
                                swapDataFixedLeg = pArray[1];
                                swapDataFloatLeg = pArray[0];
                            }
                        }
                        SwapPaymentType paymentType = swapDataFixedLeg.m_PaymentType;
                        CFixedIncomeCashFlowData[] cashFlowDataFixedLeg = null;
                        CFixedIncomeCashFlowData[] cashFlowDataFloatLeg = null;
                        if (SwapType.ZC_InflationSwap == swapType || SwapType.IRS == swapType)
                        {
                            if (FixedIncomeCashFlowData_RiskNeutral.ContainsKey(swapDataFixedLeg.m_sUniquePositionId) &&
                                FixedIncomeCashFlowData_RiskNeutral.ContainsKey(swapDataFloatLeg.m_sUniquePositionId))
                            {
                                cashFlowDataFixedLeg = new CFixedIncomeCashFlowData[2];
                                cashFlowDataFixedLeg[(int)CashFlowType.RiskNeutral] = FixedIncomeCashFlowData_RiskNeutral[swapDataFixedLeg.m_sUniquePositionId];
                                if (FixedIncomeCashFlowData_RiskRente.ContainsKey(swapDataFixedLeg.m_sUniquePositionId))
                                {
                                    cashFlowDataFixedLeg[(int)CashFlowType.RiskRente] = FixedIncomeCashFlowData_RiskRente[swapDataFixedLeg.m_sUniquePositionId];
                                }
                                else
                                {
                                    cashFlowDataFixedLeg[(int)CashFlowType.RiskRente] = new CFixedIncomeCashFlowData(swapDataFixedLeg.m_sUniquePositionId);
                                    cashFlowDataFixedLeg[(int)CashFlowType.RiskRente].finishObject();
                                }
                                cashFlowDataFloatLeg = new CFixedIncomeCashFlowData[2];
                                cashFlowDataFloatLeg[(int)CashFlowType.RiskNeutral] = FixedIncomeCashFlowData_RiskNeutral[swapDataFloatLeg.m_sUniquePositionId];
                                if (FixedIncomeCashFlowData_RiskRente.ContainsKey(swapDataFloatLeg.m_sUniquePositionId))
                                {
                                    cashFlowDataFloatLeg[(int)CashFlowType.RiskRente] = FixedIncomeCashFlowData_RiskRente[swapDataFloatLeg.m_sUniquePositionId];
                                }
                                else
                                {
                                    cashFlowDataFloatLeg[(int)CashFlowType.RiskRente] = new CFixedIncomeCashFlowData(swapDataFloatLeg.m_sUniquePositionId);
                                    cashFlowDataFloatLeg[(int)CashFlowType.RiskRente].finishObject();
                                }
                            }
                            else
                            {
                                string message = "";
                                if (SwapType.ZC_InflationSwap == swapType)
                                {
                                    message += " ZC Inflation Swap ( ";
                                    throw new Exception(message + swapDataFixedLeg.m_sSecurityID_LL + " ) has no IMW Cash flows!");
                                }
                            }
                        }
                        Instrument_Swap instrument = new Instrument_Swap(swapDataFixedLeg, swapDataFloatLeg,
                                cashFlowDataFixedLeg, cashFlowDataFloatLeg,
                                swapType, paymentType,
                                swapDataFixedLeg.m_dFrequency, swapDataFloatLeg.m_dFrequency,
                                swapDataFixedLeg.m_fFixedLegRate, swapDataFloatLeg.m_fFloatingLegRate,
                                swapDataFixedLeg.m_StartDate, swapDataFixedLeg.m_ExpiryDate.Value);
                        string ccyTranslated = Position.TranslateCurrency_Curve(scenarioZeroCurves, swapDataFixedLeg.m_sCurrency);
                        zeroCurve[0] = scenarioZeroCurves[ccyTranslated];
                        EONIASpreadCurve[0] = scenarioEONIA_SpreadCurves[ccyTranslated];
                        ccyTranslated = Position.TranslateCurrency_Curve(scenarioZeroCurves, swapDataFloatLeg.m_sCurrency);
                        zeroCurve[1] = scenarioZeroCurves[ccyTranslated];
                        EONIASpreadCurve[1] = scenarioEONIA_SpreadCurves[ccyTranslated];
                        // Create BASE scenario FX rate
                        double[] fxLevel = new double[2];
                        ccyTranslated = Position.TranslateCurrency_Fx(baseScenario, swapDataFixedLeg.m_sCurrency);
                        fxLevel[0] = baseScenario.m_Fx.ByName(ccyTranslated).m_fShockValue;
                        ccyTranslated = Position.TranslateCurrency_Fx(baseScenario, swapDataFloatLeg.m_sCurrency);
                        fxLevel[1] = baseScenario.m_Fx.ByName(ccyTranslated).m_fShockValue;
                        //                        Debug.Assert(!swapDataFixedLeg.m_sSecurityID.Equals("SW810965"));

                        try
                        {
                            IndexCPI CPIindex = null;
                            if (SwapType.ZC_InflationSwap == instrument.m_Type)
                            {
                                if (scenarioIndexCPI_List.Count > 0)
                                {
                                    CPIindex = scenarioIndexCPI_List[ccyTranslated];
                                }
                            }
                            instrument.Init(dtNow, zeroCurve, EONIASpreadCurve, CPIindex, fxLevel, swapDataFixedLeg.m_fMarketValue, swapDataFloatLeg.m_fMarketValue); // always receiver swap
                        }
                        catch (Exception exc)
                        {
                            string errorMassage = exc.Message + "; the run has stopped: Initialization of the swap (" + swapDataFixedLeg.m_sSecurityID_LL + ") has failed!";
                            errors.AddWarning(errorMassage);
                            throw new IOException(errorMassage);
                        }

                        Position position = new Position();
                        position.m_sSelectieIndex_LL = swapDataFixedLeg.m_sSelectieIndex_LL;
                        position.m_sDATA_Source = "IMW";
                        position.m_sRow = swapDataFixedLeg.m_dRow.ToString();
                        position.m_bIsLookThroughPosition = false;

                        position.m_Instrument = instrument;
                        position.m_sBalanceType = "Assets";
                        position.m_sGroup = "Swaps";
                        position.m_sSecurityType_LL = swapDataFixedLeg.m_sSecurityType;
                        position.m_sScope3 = swapDataFixedLeg.m_sScope3;
                        position.m_sScope3 = ScopeData.getScopeFormated(position.m_sScope3);
                        position.m_sSecurityID = swapDataFixedLeg.m_sSecurityID;
                        position.m_sSecurityID_LL = swapDataFixedLeg.m_sSecurityID_LL;
                        position.m_sSecurityName_LL = swapDataFixedLeg.m_sSecurityName;
                        position.m_sPortfolioId = swapDataFixedLeg.m_sPortfolioID;
                        position.m_sPositionId = position.m_sSecurityID_LL;
                        position.m_fFxRate = swapDataFixedLeg.m_fFxRate;
                        position.m_fVolume = swapDataFixedLeg.m_fNominal;
                        position.m_fAccruedDividend_LL = 0;
                        position.m_fAccruedInterest_LL = swapDataFixedLeg.m_fAccruedInterest_LL + swapDataFloatLeg.m_fAccruedInterest_LL;
                        position.m_fSCR_weight = Math.Max(0, Math.Min(1, instrument.getMaturity(dtNow)));
                        if (SwapType.FX_FORWARD == swapType || SwapType.CurrencySwap == swapType)
                        {
                            position.m_fSCR_weight = 1; // a.s.r. has a beleid to roll these instruments 
                        }
                        position.m_sCIC = swapDataFixedLeg.m_sCIC_ID;
                        position.m_sCIC_LL = swapDataFixedLeg.m_sCIC_ID_LL;
                        position.m_sCIC_SCR = position.m_sCIC_LL;
                        position.m_sAccount = swapDataFixedLeg.m_sAccount;
                        position.m_sAccount_LL = swapDataFixedLeg.m_sAccount_LL;
                        position.m_sECAP_Category_LL = swapDataFixedLeg.m_sECAP_Category_LL;

                        positions.AddPosition(position);

                        System.Windows.Forms.Application.DoEvents();
                    }
                }
            }
            catch (Exception exc)
            {
                errors.AddError("Swap legs information is loaded but Swap position (" + n + ") is not created " +
                    "(IMW row1: " + IMW_Row1 + " and IMW_Row2: " + IMW_Row2 + "): " + exc.Message);
                return new PositionList();
            }

            return positions;
        }
        // Ad-Hoc Assets:
        public PositionList Read_Ad_Hoc_Positions_Assets(DateTime dtNow, string fileName,
            ScenarioList scenarios, string sSecurityTypeRequired, bool hullWhiteModel, ErrorList errors)
        {
            const int RowStart = 2;
            // Defne the base curves:
            TotalRisk.Utilities.Scenario baseScenario = scenarios.getScenarioFairValue();
            CurveList scenarioZeroCurves = new CurveList();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
            {
                string ccy = scenarioCurve.m_sName;
                scenarioZeroCurves.Add(ccy.ToUpper(), scenarioCurve.m_Curve);
            }
            CurveList scenarioEONIA_SpreadCurves = new CurveList();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_EONIA_SpreadCurves)
            {
                string ccy = scenarioCurve.m_sName;
                scenarioEONIA_SpreadCurves.Add(ccy.ToUpper(), scenarioCurve.m_Curve);
            }
            CurveList scenarioZeroInflationCurves = new CurveList();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_InflationCurves)
            {
                string ccy = scenarioCurve.m_sName;
                scenarioZeroInflationCurves.Add(ccy.ToUpper(), scenarioCurve.m_Curve);
            }
            SortedList<string, IndexCPI> scenarioIndexCPI_List = new SortedList<string, IndexCPI>();
            foreach (ScenarioCurve scenarioCurve in baseScenario.m_InflationCurves)
            {
                string ccy = scenarioCurve.m_sName;
                scenarioZeroInflationCurves.Add(ccy.ToUpper(), scenarioCurve.m_Curve);
                IndexCPI CPIobject = new IndexCPI();
                CPIobject.SetInflationInstance(dtNow, 100, scenarioCurve.m_Curve, 100);
                scenarioIndexCPI_List.Add(ccy.ToUpper(), CPIobject);
            }
            Curve[] zeroCurve = new Curve[2];
            Curve[] EONIASpreadCurve = new Curve[2];
            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, "Assets Data", "A1");
            Dictionary<string, int> colNames = HeaderNamesColumns(values);
            PositionList positions = new PositionList();
            int row = 0;
            try
            {
                for (row = RowStart; row <= values.GetUpperBound(DimensionRow); row++)
                {
                    string securityType = ReadFieldAsString(values, row, 0, colNames, "Security Type").ToUpper();
                    if (sSecurityTypeRequired != securityType)
                    {
                        continue;
                    }
                    if ("SWAPTION" == securityType)
                    {
                        string cic_ID = ReadFieldAsString(values, row, 0, colNames, "Cic Id Ll").ToUpper();
                        string cic_ID_last2 = cic_ID.Substring(2, 2);
                        Instrument_Swaption_OriginalData swaptionData = new Instrument_Swaption_OriginalData();
                        swaptionData.m_sSelectieIndex_LL = "0000000000001000"; // 13 digit for counterparty risk
                        swaptionData.m_sScope3 = ReadFieldAsString(values, row, 0, colNames, "Scope");
                        swaptionData.m_sCurrency = ReadFieldAsString(values, row, 0, colNames, "Currency");
                        swaptionData.m_sPortfolioID = "";
                        swaptionData.m_sCICLL = cic_ID;
                        swaptionData.m_fMarketValue = ReadFieldAsDouble(values, row, 0, colNames, "Market Value");
                        swaptionData.m_fFxRate = 1.0;
                        swaptionData.m_fNominal = ReadFieldAsDouble(values, row, 0, colNames, "Nominal");
                        swaptionData.m_sType = ReadFieldAsString(values, row, 0, colNames, "Payer/Receiver");
                        swaptionData.m_fStrike = ReadFieldAsDouble(values, row, 0, colNames, "Strike");
                        DateTime? dtExpiry = ReadFieldAsDateTime(values, row, 0, colNames, "Expiry"); // swaption Expiry
                        DateTime? dtMaturity = ReadFieldAsDateTime(values, row, 0, colNames, "Maturity"); // Swap matirity
                        swaptionData.m_sInstrumentType = securityType;
                        swaptionData.m_sSecurityID = ReadFieldAsString(values, row, 0, colNames, "Security ID");
                        swaptionData.m_sSecurityName = swaptionData.m_sSecurityID;
                        SwaptionType type = cic_ID_last2.StartsWith("C") ? SwaptionType.Receiver : SwaptionType.Payer;
                        swaptionData.m_fSwaptionVolatility = ReadFieldAsDouble(values, row, 0, colNames, "Volatility");
                        // credit quality steps Group Vounterparty:
                        swaptionData.m_sGroupCounterpartyName = ReadFieldAsString(values, row, 0, colNames, "Counterprty");
                        swaptionData.m_sGroupCounterpartyLEI = swaptionData.m_sGroupCounterpartyName;
                        swaptionData.m_sGroupCounterpartyCQS = ReadFieldAsString(values, row, 0, colNames, "CQS");
                        if (swaptionData.m_sGroupCounterpartyCQS == "")
                        {//  seven as NR
                            swaptionData.m_dGroupCounterpartyCQS = 7;
                        }
                        else
                        {
                            swaptionData.m_dGroupCounterpartyCQS = Convert.ToInt32(swaptionData.m_sGroupCounterpartyCQS.Substring(0, 1));
                        }
                        bool cashSettled = true; // FM approved temporaly solution 2015-1-8
                        Instrument_Swaption instrument = new Instrument_Swaption(swaptionData, type, cashSettled, swaptionData.m_fStrike, dtExpiry.Value, dtMaturity.Value);
                        double marketValueNormalized = 0;
                        if (swaptionData.m_fNominal != 0)
                        {
                            marketValueNormalized = swaptionData.m_fMarketValue / swaptionData.m_fNominal;
                        }
                        string ccyTranslated = Position.TranslateCurrency_Curve(scenarioZeroCurves, swaptionData.m_sCurrency);
                        zeroCurve[0] = scenarioZeroCurves[ccyTranslated];
                        instrument.Init_Ad_Hoc(dtNow, zeroCurve[0], marketValueNormalized, swaptionData.m_fSwaptionVolatility);
                        Position position = new Position();
                        position.m_sSelectieIndex_LL = swaptionData.m_sSelectieIndex_LL;
                        position.m_Instrument = instrument;
                        position.m_sBalanceType = "Assets";
                        position.m_sGroup = "Swaptions";
                        position.m_sSecurityType_LL = swaptionData.m_sInstrumentType;
                        position.m_sCIC_LL = swaptionData.m_sCICLL;
                        position.m_sCIC_SCR = position.m_sCIC_LL;
                        position.m_sScope3 = swaptionData.m_sScope3;
                        position.m_sScope3 = ScopeData.getScopeFormated(position.m_sScope3);
                        position.m_sSecurityID = swaptionData.m_sSecurityID;
                        position.m_sSecurityID_LL = swaptionData.m_sSecurityID;
                        position.m_sPortfolioId = swaptionData.m_sPortfolioID;
                        position.m_sPositionId = position.m_sSecurityID_LL;
                        position.m_fFxRate = swaptionData.m_fFxRate;
                        position.m_fVolume = swaptionData.m_fNominal;
                        position.m_fSCR_weight = Math.Max(0, Math.Min(1, instrument.getMaturity(dtNow)));

                        position.m_sSecurityID = swaptionData.m_sSecurityID;
                        position.m_sCIC = swaptionData.m_sCICLL;
                        position.m_fFairValue = swaptionData.m_fMarketValue;
                        position.m_fAccruedInterest_LL = 0;
                        position.m_fAccruedDividend_LL = 0;
                        position.m_sAccount = "";
                        position.m_sAccount_LL = "";
                        position.m_sECAP_Category_LL = "";
                        position.m_sDATA_Source = "Ad_Hoc Assets";
                        positions.AddPosition(position);
                    }
                    else if ("SWAP" == securityType)
                    {
                        string cic_ID = ReadFieldAsString(values, row, 0, colNames, "Cic Id Ll").ToUpper();
                        string cic_ID_last2 = cic_ID.Substring(2, 2);
                        SwapType swapType;
                        if (cic_ID_last2 == "D1")
                        {
                            swapType = SwapType.IRS;
                        }
                        else
                        {
                            continue;
                        }
                        int dTenor = ReadFieldAsInt(values, row, 0, colNames, "Tenor");
                        double fEONIA_SPread = ReadFieldAsDouble(values, row, 0, colNames, "EONIA Spread");
                        DateTime? dtExpiry = ReadFieldAsDateTime(values, row, 0, colNames, "Expiry"); // not used for swap
                        DateTime? dtMaturity = ReadFieldAsDateTime(values, row, 0, colNames, "Maturity");
                        DateTime? dtStartDate = ReadFieldAsDateTime(values, row, 0, colNames, "Start date");
                        string sSecurity_ID = ReadFieldAsString(values, row, 0, colNames, "Security ID");
                        string sScope = ReadFieldAsString(values, row, 0, colNames, "Scope");
                        string sCurrency_FixedLeg = ReadFieldAsString(values, row, 0, colNames, "Currency");
                        // Create BASE scenario curves:
                        string sCurrency_FloatLeg = sCurrency_FixedLeg;
                        string ccyTranslated = Position.TranslateCurrency_Curve(scenarioZeroCurves, sCurrency_FixedLeg);
                        zeroCurve[0] = scenarioZeroCurves[ccyTranslated];
                        EONIASpreadCurve[0] = scenarioEONIA_SpreadCurves[ccyTranslated];
                        ccyTranslated = Position.TranslateCurrency_Curve(scenarioZeroCurves, sCurrency_FloatLeg);
                        zeroCurve[1] = scenarioZeroCurves[ccyTranslated];
                        EONIASpreadCurve[1] = scenarioEONIA_SpreadCurves[ccyTranslated];
                        double ForwardSwapRate = getSwapForwardRate(dtNow, dTenor, dtStartDate.Value, zeroCurve[0]);
                        // Create BASE scenario FX rate
                        double[] fxLevel = new double[2];
                        ccyTranslated = Position.TranslateCurrency_Fx(baseScenario, sCurrency_FixedLeg);
                        fxLevel[0] = baseScenario.m_Fx.ByName(ccyTranslated).m_fShockValue;
                        ccyTranslated = Position.TranslateCurrency_Fx(baseScenario, sCurrency_FloatLeg);
                        fxLevel[1] = baseScenario.m_Fx.ByName(ccyTranslated).m_fShockValue;

                        // Fixed Leg:
                        Instrument_Swap_OriginalData swapDataFixedLeg = new Instrument_Swap_OriginalData();
                        swapDataFixedLeg.m_dLegID = 0;
                        swapDataFixedLeg.m_dFrequency = 1;
                        swapDataFixedLeg.m_sSwapType = swapType;
                        swapDataFixedLeg.m_sSelectieIndex_LL = "0000000000001000"; // 13 digit for counterparty risk
                        swapDataFixedLeg.m_sSecurityType = securityType;
                        swapDataFixedLeg.m_sSecurityID_LL = sSecurity_ID;
                        swapDataFixedLeg.m_sSecurityName = swapDataFixedLeg.m_sSecurityID_LL;
                        swapDataFixedLeg.m_sCIC_ID_LL = cic_ID;
                        swapDataFixedLeg.m_sScope3 = sScope;
                        swapDataFixedLeg.m_fFxRate = 1.0;
                        swapDataFixedLeg.m_sCurrency = "EUR";
                        swapDataFixedLeg.m_bFixedLeg = true;
                        swapDataFixedLeg.m_ExpiryDate = dtMaturity;
                        swapDataFixedLeg.m_StartDate = dtStartDate;
                        swapDataFixedLeg.m_fNominal = ReadFieldAsDouble(values, row, 0, colNames, "Nominal");
                        swapDataFixedLeg.m_PaymentType = swapDataFixedLeg.m_fNominal > 0 ? SwapPaymentType.Receiver : SwapPaymentType.Payer;
                        swapDataFixedLeg.m_fFixedLegRate = ReadFieldAsDouble(values, row, 0, colNames, "Strike");
                        swapDataFixedLeg.m_fFixedLegRate = ForwardSwapRate;
                        // credit quality steps Group Vounterparty:
                        swapDataFixedLeg.m_sGroupCounterpartyName = ReadFieldAsString(values, row, 0, colNames, "Counterprty");
                        swapDataFixedLeg.m_sGroupCounterpartyLEI = swapDataFixedLeg.m_sGroupCounterpartyName;
                        swapDataFixedLeg.m_sGroupCounterpartyCQS = ReadFieldAsString(values, row, 0, colNames, "CQS");
                        if (swapDataFixedLeg.m_sGroupCounterpartyCQS == "")
                        {//  seven as NR
                            swapDataFixedLeg.m_dGroupCounterpartyCQS = 7;
                        }
                        else
                        {
                            swapDataFixedLeg.m_dGroupCounterpartyCQS = Convert.ToInt32(swapDataFixedLeg.m_sGroupCounterpartyCQS.Substring(0, 1));
                        }
                        // Float Leg:
                        Instrument_Swap_OriginalData swapDataFloatLeg = new Instrument_Swap_OriginalData();
                        swapDataFloatLeg.m_dLegID = 1;
                        swapDataFloatLeg.m_dFrequency = 2;
                        swapDataFloatLeg.m_fFloatingLegRate = 0;
                        swapDataFloatLeg.m_sSwapType = swapDataFixedLeg.m_sSwapType;
                        swapDataFloatLeg.m_sSelectieIndex_LL = swapDataFixedLeg.m_sSelectieIndex_LL;
                        swapDataFloatLeg.m_sSecurityType = swapDataFixedLeg.m_sSecurityType;
                        swapDataFloatLeg.m_sSecurityName = swapDataFixedLeg.m_sSecurityID_LL;
                        swapDataFloatLeg.m_sCIC_ID_LL = cic_ID;
                        swapDataFloatLeg.m_sScope3 = sScope;
                        swapDataFloatLeg.m_fFxRate = 1.0;
                        swapDataFloatLeg.m_sCurrency = "EUR";
                        swapDataFloatLeg.m_bFixedLeg = false;
                        swapDataFloatLeg.m_ExpiryDate = dtMaturity;
                        swapDataFloatLeg.m_StartDate = dtStartDate;
                        swapDataFloatLeg.m_fNominal = -swapDataFixedLeg.m_fNominal;
                        swapDataFloatLeg.m_PaymentType = swapDataFixedLeg.m_PaymentType;
                        swapDataFloatLeg.m_sGroupCounterpartyName = swapDataFixedLeg.m_sGroupCounterpartyName;
                        swapDataFloatLeg.m_sGroupCounterpartyLEI = swapDataFixedLeg.m_sGroupCounterpartyLEI;
                        swapDataFloatLeg.m_sGroupCounterpartyCQS = swapDataFixedLeg.m_sGroupCounterpartyCQS;
                        swapDataFloatLeg.m_dGroupCounterpartyCQS = swapDataFixedLeg.m_dGroupCounterpartyCQS;
                        // create the swap:
                        Instrument_Swap instrument = new Instrument_Swap(swapDataFixedLeg, swapDataFloatLeg,
                                swapType, swapDataFixedLeg.m_PaymentType,
                                swapDataFixedLeg.m_dFrequency, swapDataFloatLeg.m_dFrequency,
                                swapDataFixedLeg.m_fFixedLegRate, swapDataFloatLeg.m_fFloatingLegRate,
                                swapDataFixedLeg.m_StartDate, swapDataFixedLeg.m_ExpiryDate.Value);

                        try
                        {
                            IndexCPI CPIindex = null;
                            if (SwapType.ZC_InflationSwap == instrument.m_Type)
                            {
                                if (scenarioIndexCPI_List.Count > 0)
                                {
                                    CPIindex = scenarioIndexCPI_List[ccyTranslated];
                                }
                            }
                            instrument.Init_Ad_Hoc(dtNow, zeroCurve, EONIASpreadCurve, CPIindex, fxLevel, swapDataFixedLeg.m_fMarketValue, swapDataFloatLeg.m_fMarketValue, fEONIA_SPread); // always receiver swap
                        }
                        catch (Exception exc)
                        {
                            string errorMassage = exc.Message + "; the run has stopped: Initialization of the ad-hoc swap (" + swapDataFixedLeg.m_sSecurityID_LL + ") has failed!";
                            errors.AddWarning(errorMassage);
                            throw new IOException(errorMassage);
                        }
                        // Position:
                        Position position = new Position();
                        position.m_sSelectieIndex_LL = swapDataFixedLeg.m_sSelectieIndex_LL;
                        position.m_Instrument = instrument;
                        position.m_sBalanceType = "Assets";
                        position.m_sGroup = "Swaps";
                        position.m_sSecurityType_LL = swapDataFixedLeg.m_sSecurityType;
                        position.m_sScope3 = swapDataFixedLeg.m_sScope3;
                        position.m_sScope3 = ScopeData.getScopeFormated(position.m_sScope3);
                        position.m_sSecurityID = swapDataFixedLeg.m_sSecurityID;
                        position.m_sSecurityID_LL = swapDataFixedLeg.m_sSecurityID_LL;
                        position.m_sSecurityName_LL = swapDataFixedLeg.m_sSecurityName;
                        position.m_sPortfolioId = position.m_sScope3;
                        position.m_sPositionId = position.m_sSecurityID_LL;
                        position.m_fFxRate = swapDataFixedLeg.m_fFxRate;
                        position.m_fVolume = swapDataFixedLeg.m_fNominal;
                        position.m_fAccruedDividend_LL = 0;
                        position.m_fAccruedInterest_LL = swapDataFixedLeg.m_fAccruedInterest_LL + swapDataFloatLeg.m_fAccruedInterest_LL;
                        position.m_fSCR_weight = Math.Max(0, Math.Min(1, instrument.getMaturity(dtNow)));

                        position.m_sCIC = swapDataFixedLeg.m_sCIC_ID;
                        position.m_sCIC_LL = swapDataFixedLeg.m_sCIC_ID_LL;
                        position.m_sCIC_SCR = position.m_sCIC_LL;
                        position.m_sAccount = swapDataFixedLeg.m_sAccount;
                        position.m_sAccount_LL = swapDataFixedLeg.m_sAccount_LL;
                        position.m_sECAP_Category_LL = swapDataFixedLeg.m_sECAP_Category_LL;

                        position.m_sDATA_Source = "Ad_Hoc Assets";
                        positions.AddPosition(position);
                    }
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            catch (Exception exc)
            {
                errors.AddError("Fout tijdens inlezen swaption positie in rij " + row + " : " + exc.Message);
            }

            return positions;
        }
        public double getSwapForwardRate(DateTime dtNow, int tenor, DateTime startDate, Curve zeroCurve)
        {
            double start = DateTimeExtensions.YearFrac(dtNow, startDate, Daycount.ACT_ACT);
            double sumDf = 0;
            double[] dfSwap = new double[2 * tenor + 1];
            for (int idx = 0; idx < dfSwap.Length; idx++)
            {
                dfSwap[idx] = zeroCurve.DiscountFactor(start + 0.5 * idx);
                if (idx > 0)
                {
                    sumDf += dfSwap[idx] * 0.5;
                }
            }
            double fwd_rate = (dfSwap[0] - dfSwap[dfSwap.Length - 1]) / sumDf;
            return fwd_rate;
        }

        // Real Estate:
        public SortedList<string, string> ReadScopeMappingFromConsolidationFile(string fileName, ErrorList errors)
        {
            SortedList<string, string> maping = new SortedList<string, string>();
            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, "", "A1");
            Dictionary<string, int> headers = HeaderNamesColumns(values); // import headers
            for (int row = values.GetLowerBound(DimensionRow) + 1; row <= values.GetUpperBound(DimensionRow); row++)
            {
                string sSAP_Entity = ReadFieldAsString(values, row, 0, headers, "SAP ENTITY").Trim();
                string sScope3 = ReadFieldAsString(values, row, 0, headers, "CONS_ECAP").Trim();
                if (maping.ContainsKey(sSAP_Entity))
                {
                    if (sScope3 != maping[sSAP_Entity])
                    {
                        string message = "SAP Entity cannot have more than 1 Scope3 code, row " + row.ToString();
                        if (MessageBox.Show(message + ". Bestand alsnog verwerken?", "Ongeldige Scope3 code", MessageBoxButtons.OKCancel) == DialogResult.OK)
                        {
                            if (errors != null) errors.AddWarning("Ongeldige SAP Entity in bestand. Gebruiker heeft melding genegeerd");
                        }
                        else
                        {
                            if (errors != null) errors.AddError("Ongeldige SAP Entity in bestand. Gebruiker heeft verwerking gestopt");
                            return new SortedList<string, string>();
                        }
                    }
                }
                else
                {
                    maping.Add(sSAP_Entity, sScope3);
                }
            }

            return maping;
        }

        public PositionList ReadRealEstatePositions_IMW(DateTime dtNow, string IMW_FileName, ErrorList errors)
        {
            PositionList positions = new PositionList();

            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(IMW_FileName, "Report 1", "A1");
            // Read column names into dictionary to map column name to column number
            Dictionary<string, int> columnNames = HeaderNamesColumns(values);

            for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
            {
                Instrument_RealEstate_OriginalData entry = new Instrument_RealEstate_OriginalData();
                entry.m_sBalanceType = "Assets";
                entry.m_sGroup = "Real Estate";
                entry.m_dtReport = ReadFieldAsDateTime2(values, row, 0, columnNames, "Reporting Date");
                entry.m_sDataSource = ReadFieldAsString(values, row, 0, columnNames, "Source Laagste Lt Level");
                entry.m_sSAP_Entity_ID = ReadFieldAsString(values, row, 0, columnNames, "SmS Entity Id");
                entry.m_sTagetikAccount = ReadFieldAsString(values, row, 0, columnNames, "Tagetik Account");
                entry.m_sTagetikAccountName = ReadFieldAsString(values, row, 0, columnNames, "Tagetik Account name");
                entry.m_sTagetikAccount_LL = ReadFieldAsString(values, row, 0, columnNames, "Tagetik Account (Look through)");
                entry.m_sTagetikAccountName_LL = ReadFieldAsString(values, row, 0, columnNames, "Tagetik Account name (Look through)");
                entry.m_sInfrastructInvestmentID = ReadFieldAsString(values, row, 0, columnNames, "Infrastruct Investment Id");
                entry.m_sScope3 = ReadFieldAsString(values, row, 0, columnNames, "Ecs Cons Ecap");
                entry.m_sScope3 = ScopeData.getScopeFormated(entry.m_sScope3);

                entry.m_sCIC_ID = ReadFieldAsString(values, row, 0, columnNames, "Cic Id");
                entry.m_sCIC_ID_LL = ReadFieldAsString(values, row, 0, columnNames, "Cic Id Laagste Lt Level");
                entry.m_sSecurity_ID = ReadFieldAsString(values, row, 0, columnNames, "Security Id");
                entry.m_sSecurity_ID_LL = ReadFieldAsString(values, row, 0, columnNames, "Security Id Laagste Lt Level");
                entry.m_sSecurity_Name = ReadFieldAsString(values, row, 0, columnNames, "Security Name");
                entry.m_sSecurity_Name_LL = ReadFieldAsString(values, row, 0, columnNames, "Security Name Laagste Lt Level");
                entry.m_sSecurity_Type_LL = ReadFieldAsString(values, row, 0, columnNames, "Security Type Ll");
                entry.m_sECAP_AllocationType = ReadFieldAsString(values, row, 0, columnNames, "ECAP Allocation Type");
                entry.m_sCountryCode = ReadFieldAsString(values, row, 0, columnNames, "Country Code Laagste Lt Level");
                entry.m_sCurrency = ReadFieldAsString(values, row, 0, columnNames, "Currency Laagste Lt Level");
                entry.m_fBalcostval_Pc_LL = ReadFieldAsDouble(values, row, 0, columnNames, "Balcostval Pc Laagste Lt Level");
                entry.m_MarketValue = ReadFieldAsDouble(values, row, 0, columnNames, "Market Value Eur Ll");
                entry.m_bICO = ReadFieldAsBool(values, row, 0, columnNames, "Eliminatie ASR LL");

                Instrument_RealEstate instrument = new Instrument_RealEstate(entry);
                Position position = new Position();
                position.m_sDATA_Source = entry.m_sDataSource;
                position.m_sRow = row.ToString();
                position.m_bIsLookThroughPosition = true;

                position.m_Instrument = instrument;
                position.m_sBalanceType = entry.m_sBalanceType;
                position.m_sGroup = entry.m_sGroup;
                position.m_sScope3 = entry.m_sScope3;
                position.m_bICO = entry.m_bICO;
                position.m_fVolume = entry.m_fBalcostval_Pc_LL;
                position.m_fFairValue = entry.m_MarketValue;
                position.m_sSecurityType_LL = "REAL ESTATE";
                position.m_sCIC = entry.m_sCIC_ID;
                position.m_sCIC_LL = entry.m_sCIC_ID_LL;
                position.m_sCIC_SCR = position.m_sCIC_LL;
                position.m_sPortfolioId = entry.m_sSAP_Entity_ID;
                position.m_sPositionId = entry.m_sSecurity_ID_LL;
                position.m_sSecurityID_LL = entry.m_sSecurity_ID_LL;
                position.m_sSecurityName_LL = entry.m_sSecurity_Name_LL;
                position.m_sCurrency = entry.m_sCurrency;
                position.m_sCountryCurrency = entry.m_sCurrency;
                position.m_sDATA_Source = entry.m_sDataSource;
                position.m_sECAP_Category_LL = instrument.m_OriginalData.m_sECAP_AllocationType;
                position.m_sSecurityCreditQuality = "7";
                positions.AddPosition(position);
            }

            return positions;
        }
        public PositionList ReadRealEstatePositions_Saldi(DateTime dtNow, SortedList<string, string> SAP_Entity_vs_Scope3_mapping, string Saldi_FileName, ErrorList errors)
        {

            Dictionary<string, string> TAGETIK_accounts_vs_CIC = new Dictionary<string, string>();
            TAGETIK_accounts_vs_CIC.Add("A102101010", "XT99");
            TAGETIK_accounts_vs_CIC.Add("A102101020", "XT99");
            TAGETIK_accounts_vs_CIC.Add("A102101040", "XT99");

            TAGETIK_accounts_vs_CIC.Add("A102201010", "XT99");
            TAGETIK_accounts_vs_CIC.Add("A102201020", "XT99");

            TAGETIK_accounts_vs_CIC.Add("A101201010", "XT91");
            TAGETIK_accounts_vs_CIC.Add("A101201110", "XT91");
            TAGETIK_accounts_vs_CIC.Add("A101101020", "XT91");
            TAGETIK_accounts_vs_CIC.Add("A101101120", "XT91");
            TAGETIK_accounts_vs_CIC.Add("A101201230", "XT91");

            Dictionary<string, string> TAGETIK_accounts_vs_Category = new Dictionary<string, string>();
            TAGETIK_accounts_vs_Category.Add("A102101010", "Other");
            TAGETIK_accounts_vs_Category.Add("A102101020", "Other");
            TAGETIK_accounts_vs_Category.Add("A102101040", "Other");

            TAGETIK_accounts_vs_Category.Add("A102201010", "Other");
            TAGETIK_accounts_vs_Category.Add("A102201020", "Other");

            TAGETIK_accounts_vs_Category.Add("A101201010", "Offices");
            TAGETIK_accounts_vs_Category.Add("A101201110", "Offices");
            TAGETIK_accounts_vs_Category.Add("A101101020", "Offices");
            TAGETIK_accounts_vs_Category.Add("A101101120", "Offices");
            TAGETIK_accounts_vs_Category.Add("A101201230", "Offices");

            Dictionary<string, List<string>> Offices_SAP_Entities_Excluded = new Dictionary<string, List<string>>();
            Offices_SAP_Entities_Excluded["A101201010"] = new List<string>();
            Offices_SAP_Entities_Excluded["A101201110"] = new List<string>();
            Offices_SAP_Entities_Excluded["A101101020"] = new List<string>();
            Offices_SAP_Entities_Excluded["A101101120"] = new List<string>();
            Offices_SAP_Entities_Excluded["A101201230"] = new List<string>();
            foreach (KeyValuePair<string, List<string>> account in Offices_SAP_Entities_Excluded)
            {
                Offices_SAP_Entities_Excluded[account.Key].Add("VLS");
            }
            PositionList positions = new PositionList();

            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(Saldi_FileName, "", "A3");
            // Read column names into dictionary to map column name to column number
            Dictionary<string, int> columnNames = HeaderNamesColumns(values);

            for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
            {
                string sTAGETIK_Account = ReadFieldAsString(values, row, 0, columnNames, "TGK ACCOUNT").Trim();
                string sSAP_Account = ReadFieldAsString(values, row, 0, columnNames, "SAP ACCOUNT").ToUpper(); // 07260; 06560 to be removed from SaldiTabel data
                string sSAP_Entity = ReadFieldAsString(values, row, 0, columnNames, "SAP ENTITY").Trim().ToUpper();
                if (!TAGETIK_accounts_vs_Category.ContainsKey(sTAGETIK_Account))
                {
                    continue;
                }
                if (Offices_SAP_Entities_Excluded.ContainsKey(sTAGETIK_Account) &&
                    Offices_SAP_Entities_Excluded[sTAGETIK_Account].Contains(sSAP_Entity))
                {
                    continue;
                }

                string sTAGETIK_Account_Desc = ReadFieldAsString(values, row, 0, columnNames, "TGK ACCOUNT DESC").Trim();
                string sSAP_Account_Desc = ReadFieldAsString(values, row, 0, columnNames, "SAP ACCOUNT DESC").ToUpper();

                Instrument_RealEstate_OriginalData entry = new Instrument_RealEstate_OriginalData();
                entry.m_sBalanceType = "Assets";
                entry.m_sGroup = "Real Estate";
                string test = ReadFieldAsString(values, row, 0, columnNames, "SCENARIO");
                /*
                int year = 2018;
                int.TryParse(test.Substring(0, 4), out year); 
                int month = ReadFieldAsInt(values, row, 0, columnNames, "PERIOD");
                int day = DateTime.DaysInMonth(year, month);
                entry.m_dtReport = new DateTime(year,month,day);
                */
                entry.m_dtReport = ReadFieldAsDateTime(values, row, 0, columnNames, "PERIOD");
                entry.m_sDataSource = "SaldiTabel";
                entry.m_sSAP_Entity_ID = sSAP_Entity;
                entry.m_sTagetikAccount = sTAGETIK_Account;
                entry.m_sTagetikAccountName = sTAGETIK_Account_Desc;
                entry.m_sTagetikAccount_LL = sSAP_Account;
                entry.m_sTagetikAccountName_LL = sSAP_Account_Desc;
                entry.m_sInfrastructInvestmentID = "";
                entry.m_sScope3 = ScopeData.getScopeFormated(SAP_Entity_vs_Scope3_mapping[sSAP_Entity]);

                entry.m_sCIC_ID = TAGETIK_accounts_vs_CIC[sTAGETIK_Account];
                entry.m_sCIC_ID_LL = entry.m_sCIC_ID;
                entry.m_sSecurity_ID = sTAGETIK_Account;
                entry.m_sSecurity_ID_LL = sSAP_Account;
                entry.m_sSecurity_Name = sTAGETIK_Account_Desc;
                entry.m_sSecurity_Name_LL = sSAP_Account_Desc;
                entry.m_sSecurity_Type_LL = "REAL ESTATE";
                entry.m_sECAP_AllocationType = TAGETIK_accounts_vs_Category[sTAGETIK_Account];
                entry.m_sCountryCode = "NDL";
                entry.m_sCurrency = "EUR";
                entry.m_fBalcostval_Pc_LL = ReadFieldAsDouble(values, row, 0, columnNames, "SAP SALDO ORG");
                entry.m_MarketValue = entry.m_fBalcostval_Pc_LL;
                test = ReadFieldAsString(values, row, 0, columnNames, "SAP PARTNER");
                if ("" != test)
                {
                    entry.m_bICO = true;
                }
                else
                {
                    entry.m_bICO = false;
                }

                Instrument_RealEstate instrument = new Instrument_RealEstate(entry);
                Position position = new Position();
                position.m_sDATA_Source = entry.m_sDataSource;
                position.m_sRow = row.ToString();
                position.m_bIsLookThroughPosition = false;

                position.m_Instrument = instrument;
                position.m_sBalanceType = entry.m_sBalanceType;
                position.m_sGroup = entry.m_sGroup;
                position.m_sScope3 = entry.m_sScope3;
                position.m_fVolume = entry.m_fBalcostval_Pc_LL;
                position.m_fFairValue = entry.m_MarketValue;
                position.m_sSecurityType_LL = "REAL ESTATE";
                position.m_sCIC = entry.m_sCIC_ID;
                position.m_sCIC_LL = entry.m_sCIC_ID_LL;
                position.m_sCIC_SCR = position.m_sCIC_LL;
                position.m_sPortfolioId = entry.m_sSAP_Entity_ID;
                position.m_sPositionId = "TGK_" + sTAGETIK_Account + "_SAP_" + sSAP_Account;
                position.m_sSecurityID_LL = entry.m_sSecurity_ID_LL;
                position.m_sSecurityName_LL = entry.m_sSecurity_Name_LL;
                position.m_sCurrency = entry.m_sCurrency;
                position.m_sCountryCurrency = entry.m_sCurrency;
                position.m_sDATA_Source = entry.m_sDataSource;
                position.m_sECAP_Category_LL = instrument.m_OriginalData.m_sECAP_AllocationType;
                position.m_sSecurityCreditQuality = "7";
                positions.AddPosition(position);
            }

            return positions;
        }
        // Liabilities:
        public PositionList ReadOSMModelPoints(DateTime dtNow, string fileName, ErrorList errors, ScopeData scopeData)
        {
            PositionList positions = new PositionList();
            List<string> InstrumentTypes = new List<string>();
            InstrumentTypes.Add("Liabilities");
            InstrumentTypes.Add("UL value");
            if (fileName != "")
            {
                object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, "MW Liabilities", "A1");
                Dictionary<string, int> headers = HeaderNamesColumns(values); // import headers
                for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
                {
                    string sInstrumentType = ReadFieldAsString(values, row, 0, headers, "strInvestmentType");
                    if (InstrumentTypes.Contains(sInstrumentType))
                    {
                        Instrument_OSM_OriginalData entry = new Instrument_OSM_OriginalData();
                        entry.m_sDataSource = "OSM";
                        entry.m_sRow = row.ToString();
                        entry.m_sScope3 = ReadFieldAsString(values, row, 0, headers, "strScope");
                        entry.m_sScope3 = ScopeData.getScopeFormated(entry.m_sScope3);
                        entry.m_sInstrumentType = sInstrumentType;
                        if ("UL value" == sInstrumentType)
                        {
                            entry.m_sGroup = "Unit Linked";
                        }
                        else if ("Liabilities" == sInstrumentType)
                        {
                            entry.m_sGroup = "Unknown";
                            if (scopeData.m_Liabilities_Groups.ContainsKey(entry.m_sScope3))
                            {
                                entry.m_sGroup = scopeData.m_Liabilities_Groups[entry.m_sScope3];
                            }
                        }
                        entry.m_sModelPointID = ReadFieldAsString(values, row, 0, headers, "strModelPoint").ToUpper();
                        entry.m_sCategory = ReadFieldAsString(values, row, 0, headers, "strCategory").ToUpper();
                        if (HeaderNameExists(headers, "Currency"))
                        {
                            entry.m_sCurrency = ReadFieldAsString(values, row, 0, headers, "Currency").ToUpper();
                        }
                        else
                        {
                            entry.m_sCurrency = "EUR";
                        }
                        entry.m_fFxRate = 1.0 / ReadFieldAsDouble(values, row, 0, headers, "FX");
                        entry.m_fNominal = ReadFieldAsDouble(values, row, 0, headers, "Cost");
                        int firstColumnScenarios = headers["FairValue".ToUpper()];
                        for (int col = firstColumnScenarios; col <= values.GetUpperBound(DimensionCol); col++)
                        {
                            double marketValue = ReadFieldAsDouble(values, row, col);
                            string scenarioName = ReadFieldAsString(values, 1, col).Replace(" ", "_");
                            entry.m_MarketValues.Add(scenarioName, marketValue);
                        }
                        Instrument_OSM instrument = new Instrument_OSM(entry);
                        Position position = new Position();
                        position.m_sDATA_Source = entry.m_sDataSource;
                        position.m_sRow = "'" + entry.m_sRow;
                        position.m_Instrument = instrument;
                        position.m_sBalanceType = entry.m_sBalanceType;
                        position.m_sGroup = entry.m_sGroup;
                        position.m_sScope3 = instrument.m_OriginalData.m_sScope3;
                        position.m_fVolume = instrument.m_OriginalData.m_fNominal;
                        position.m_sSecurityType_LL = instrument.m_OriginalData.m_sInstrumentType;
                        position.m_sCIC_LL = "";
                        position.m_sCIC_SCR = position.m_sCIC_LL;
                        position.m_sPortfolioId = instrument.m_OriginalData.m_sModelPointID;
                        position.m_sPositionId = instrument.m_OriginalData.m_sModelPointID;
                        position.m_sSecurityID_LL = instrument.m_OriginalData.m_sModelPointID;
                        position.m_sSecurityName_LL = instrument.m_OriginalData.m_sModelPointID;
                        position.m_sDATA_Source = instrument.m_OriginalData.m_sDataSource;
                        position.m_sCurrency = entry.m_sCurrency;
                        position.m_sCountryCurrency = entry.m_sCurrency;
                        positions.AddPosition(position);
                    }
                }
            }
            return positions;
        }
        public PositionList ReadULGModelPoints(DateTime dtNow, string fileName, ErrorList errors)
        {
            PositionList positions = new PositionList();
            if (fileName != "")
            {
                object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, "Opgave voor FRM fixed", "A1");
                Dictionary<string, int> headers = HeaderNamesColumns(values); // import headers
                for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
                {
                    Instrument_ULG_OriginalData entry = new Instrument_ULG_OriginalData();
                    entry.m_sDataSource = "ULG";
                    entry.m_sRow = row.ToString();
                    entry.m_sModelPointID = ReadFieldAsString(values, row, 0, headers, "Totaal").ToUpper();
                    entry.m_sInstrumentType = "UL Guarantee";
                    entry.m_sScope3 = ReadFieldAsString(values, row, 0, headers, "Scope");
                    entry.m_sScope3 = ScopeData.getScopeFormated(entry.m_sScope3);
                    entry.m_sCurrency = "EUR";
                    entry.m_fFxRate = 1.0;
                    int firstColumnScenarios = headers["FairValue".ToUpper()];
                    for (int col = firstColumnScenarios; col <= values.GetUpperBound(DimensionCol); col++)
                    {
                        double marketValue = -ReadFieldAsDouble(values, row, col);
                        string scenarioName = ReadFieldAsString(values, 1, col).Replace(" ", "_");
                        entry.m_MarketValues.Add(scenarioName, marketValue);
                    }
                    Instrument_ULG instrument = new Instrument_ULG(entry);
                    Position position = new Position();
                    position.m_sDATA_Source = entry.m_sDataSource;
                    position.m_sRow = "'" + entry.m_sRow;
                    position.m_Instrument = instrument;
                    position.m_sBalanceType = entry.m_sBalanceType;
                    position.m_sGroup = entry.m_sGroup;
                    position.m_sScope3 = instrument.m_OriginalData.m_sScope3;
                    position.m_fVolume = 0;
                    position.m_sSecurityType_LL = instrument.m_OriginalData.m_sInstrumentType;
                    position.m_sCIC_LL = "";
                    position.m_sCIC_SCR = position.m_sCIC_LL;
                    position.m_sPortfolioId = instrument.m_OriginalData.m_sModelPointID;
                    position.m_sPositionId = instrument.m_OriginalData.m_sModelPointID;
                    position.m_sSecurityID_LL = instrument.m_OriginalData.m_sModelPointID;
                    position.m_sSecurityName_LL = instrument.m_OriginalData.m_sModelPointID;
                    position.m_sCurrency = entry.m_sCurrency;
                    position.m_sCountryCurrency = entry.m_sCurrency;
                    position.m_sDATA_Source = instrument.m_OriginalData.m_sDataSource;
                    positions.AddPosition(position);
                }
            }
            return positions;
        }
        public PositionList ReadRiskMarginPerOTSO(DateTime dtNow, string fileName, ErrorList errors)
        {
            PositionList positions = new PositionList();
            if (fileName != "")
            {
                object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, "OT (renterisico)", "A4");
                Dictionary<string, int> headers = HeaderNamesColumns(values); // import headers
                for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
                {
                    Instrument_RiskMargin_OriginalData entry = new Instrument_RiskMargin_OriginalData();
                    entry.m_sDataSource = "RiskMarginTool";
                    entry.m_sRow = row.ToString();
                    entry.m_sEntity = ReadFieldAsString(values, row, 1);
                    entry.m_sInstrumentType = entry.m_sGroup;
                    entry.m_sScope3 = "";
                    entry.m_sCurrency = "EUR";
                    entry.m_fFxRate = 1.0;
                    int firstColumnScenarios = headers["FairValue".ToUpper()];
                    for (int col = firstColumnScenarios; col <= values.GetUpperBound(DimensionCol); col++)
                    {
                        double marketValue = ReadFieldAsDouble(values, row, col);
                        string scenarioName = ReadFieldAsString(values, 1, col).Replace(" ", "_");
                        entry.m_MarketValues.Add(scenarioName, marketValue);
                    }
                    Instrument_RiskMargin instrument = new Instrument_RiskMargin(entry);
                    Position position = new Position();
                    position.m_sDATA_Source = entry.m_sDataSource;
                    position.m_sRow = "'" + entry.m_sRow;
                    position.m_Instrument = instrument;
                    position.m_sBalanceType = entry.m_sBalanceType;
                    position.m_sGroup = entry.m_sGroup;
                    position.m_sEntity = instrument.m_OriginalData.m_sEntity;
                    position.m_sScope3 = instrument.m_OriginalData.m_sScope3;
                    position.m_fVolume = 0;
                    position.m_sSecurityType_LL = instrument.m_OriginalData.m_sInstrumentType;
                    position.m_sCIC_LL = "";
                    position.m_sCIC_SCR = position.m_sCIC_LL;
                    position.m_sPortfolioId = instrument.m_OriginalData.m_sEntity;
                    position.m_sPositionId = instrument.m_OriginalData.m_sGroup;
                    position.m_sSecurityID_LL = instrument.m_OriginalData.m_sGroup;
                    position.m_sSecurityName_LL = instrument.m_OriginalData.m_sGroup;
                    position.m_sCurrency = entry.m_sCurrency;
                    position.m_sCountryCurrency = entry.m_sCurrency;
                    position.m_sDATA_Source = instrument.m_OriginalData.m_sDataSource;
                    positions.AddPosition(position);
                }
            }
            return positions;
        }
        public PositionList ReadRiskMargin_RAW(DateTime dtNow, Dictionary<string, object[,]> RiskMargin_Tabs, ErrorList errors)
        {
            RiskMargin_RAWData rawData = null;
            foreach (KeyValuePair<string, object[,]> pair in RiskMargin_Tabs)
            {
                string tabName = pair.Key.Trim().ToUpper(); // "RUN_208_DET_CF_L"  and "RUN_208_DET_CF_P";
                string name_BL;
                if (tabName.EndsWith("_L"))
                {
                    name_BL = "LEVEN";
                }
                else if (tabName.EndsWith("_P"))
                {
                    name_BL = "PENSION";
                }
                else
                {
                    name_BL = "";
                    // error unknown Business Unit: Exception
                }
                string name_SC = tabName.Substring(0, 14);
                Dictionary<string, Dictionary<string, double[]>> data_SC =
                    new Dictionary<string, Dictionary<string, double[]>>(); // [ModelPoint][Attribute][values]
                object[,] values = pair.Value;
                Dictionary<string, int> headers = HeaderNamesColumns(values); // import headers
                int firstColumnValues = headers["VAR_NAME"] + 1;

                int[] periodsInMonths = new int[headers.Count() - headers["VAR_NAME"]];
                for (int col = firstColumnValues; col <= values.GetUpperBound(DimensionCol); col++)
                {
                    periodsInMonths[col - firstColumnValues] = ReadFieldAsInt(values, 1, col);
                }
                for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
                {
                    string name_MP = ReadFieldAsString(values, row, 0, headers, "LIAB_NAME").Trim().ToUpper();
                    string name_AT = ReadFieldAsString(values, row, 0, headers, "VAR_NAME").Trim().ToUpper();
                    double[] Attributes_values = new double[periodsInMonths.Length];
                    for (int col = firstColumnValues; col <= values.GetUpperBound(DimensionCol); col++)
                    {
                        Attributes_values[col - firstColumnValues] = ReadFieldAsDouble(values, row, col);
                    }
                    if (!data_SC.ContainsKey(name_MP))
                    {
                        data_SC[name_MP] = new Dictionary<string, double[]>();
                    }
                    if (data_SC[name_MP].ContainsKey(name_AT))
                    {
                        // error: the Attribute data are filled in, Exception
                    }
                    data_SC[name_MP][name_AT] = Attributes_values;
                }
                if (null == rawData)
                {
                    rawData = RiskMargin_RAWData.getInstance(dtNow, periodsInMonths);
                }
                rawData.AddScenario(name_BL, name_SC, periodsInMonths, data_SC);
            }
            PositionList positions = rawData.getPositions();
            return positions;
        }
        public Dictionary<string, double[][]> ReadCorrelationTable(string fileName)
        {
            string sheetName = "SCR Martket Risk Correlations";
            ExcelWrapper.ExcelWrapper ExcelObj = new ExcelWrapper.ExcelWrapper();
            Workbook inputFile = null;
            Worksheet tab = null;
            try
            {
                inputFile = ExcelObj.WorkbookOpen(fileName, true);
            }
            catch
            {
                throw new IOException("Bestand '" + fileName + "' kan niet worden geopend");
            }

            try
            {
                tab = inputFile.Worksheets[sheetName];
            }
            catch
            {
                throw new IOException("Sheet '" + sheetName + "' in bestand '" + fileName + "' kan niet worden geopend");
            }
            Dictionary<string, double[][]> ReturnValue = new Dictionary<string, double[][]>();
            Range xlsRange = null;
            Range range;
            object[,] values = null;
            for (int j = 0; j < 9; j++)
            {
                string TableName = "";
                int size = 0;
                switch (j)
                {
                    case 0:
                        TableName = "RiskCorrelationDown";
                        size = 6;
                        range = ExcelWrapper.ExcelWrapper.RangeSetCurrentRegion(tab, "A1");
                        values = ExcelWrapper.ExcelWrapper.ReadRange(tab, range);
                        break;
                    case 1:
                        TableName = "RiskCorrelationUp";
                        size = 6;
                        range = ExcelWrapper.ExcelWrapper.RangeSetCurrentRegion(tab, "A9");
                        values = ExcelWrapper.ExcelWrapper.ReadRange(tab, range);
                        break;
                    case 2:
                        TableName = "RiskCorrelationECAP";
                        size = 7;
                        range = ExcelWrapper.ExcelWrapper.RangeSetCurrentRegion(tab, "A17");
                        values = ExcelWrapper.ExcelWrapper.ReadRange(tab, range);
                        break;
                    case 3:
                        TableName = "EquityCorrelationECAP";
                        size = 5;
                        range = ExcelWrapper.ExcelWrapper.RangeSetCurrentRegion(tab, "A26");
                        values = ExcelWrapper.ExcelWrapper.ReadRange(tab, range);
                        break;
                    case 4:
                        TableName = "RiskCorrelation_Life_Module_SCR";
                        size = 7;
                        range = ExcelWrapper.ExcelWrapper.RangeSetCurrentRegion(tab, "A33");
                        values = ExcelWrapper.ExcelWrapper.ReadRange(tab, range);
                        break;
                    case 5:
                        TableName = "RiskCorrelation_All_Modules_SCR";
                        size = 5;
                        range = ExcelWrapper.ExcelWrapper.RangeSetCurrentRegion(tab, "A42");
                        values = ExcelWrapper.ExcelWrapper.ReadRange(tab, range);
                        break;
                    case 6:
                        TableName = "RiskCorrelation_NonLife_Module_SCR";
                        size = 3;
                        range = ExcelWrapper.ExcelWrapper.RangeSetCurrentRegion(tab, "A49");
                        values = ExcelWrapper.ExcelWrapper.ReadRange(tab, range);
                        break;
                    case 7:
                        TableName = "RiskCorrelation_Health_Module_SCR";
                        size = 3;
                        range = ExcelWrapper.ExcelWrapper.RangeSetCurrentRegion(tab, "A54");
                        values = ExcelWrapper.ExcelWrapper.ReadRange(tab, range);
                        break;
                    case 8:
                        TableName = "RiskCorrelation_Health_SLT_Module_SCR";
                        size = 5;
                        range = ExcelWrapper.ExcelWrapper.RangeSetCurrentRegion(tab, "A59");
                        values = ExcelWrapper.ExcelWrapper.ReadRange(tab, range);
                        break;
                    default:
                        break;
                }
                Dictionary<string, int> headers = HeaderNamesColumns(values); // import headers
                size = headers.Count() - 1;
                double[][] ary = new double[size][];
                for (int i = 0; i < size; i++) ary[i] = new double[size];
                int shiftRow = values.GetLowerBound(DimensionRow) + 1;
                int shiftCol = values.GetLowerBound(DimensionCol) + 1;
                TableName = ReadFieldAsString(values, shiftRow - 1, shiftCol - 1);
                for (int row = values.GetLowerBound(DimensionRow) + 1; row <= values.GetUpperBound(DimensionRow); row++)
                {
                    for (int col = values.GetLowerBound(DimensionCol) + 1; col <= values.GetUpperBound(DimensionCol); col++)
                    {
                        ary[row - shiftRow][col - shiftCol] = ReadFieldAsDouble(values, row, col);
                    }
                }
                /*
                object[,] read_in = ExcelWrapper.ExcelWrapper.ReadRange(tab, xlsRange);
                for (int m = 0; m < size; m++)
                {
                    for (int n = 0; n < size; n++)
                    {
                        ary[m][n] = Convert.ToDouble(read_in[m + 1, n + 1]);
                    }
                }
                */
                ReturnValue[TableName] = ary;
            }
            inputFile.Close();
            ExcelObj.Dispose();
            return ReturnValue;
        }

        // Mortgages:
        public PositionList ReadMortgagePositions_mortgagePackages(DateTime reportDate, string fileName, string sheetName,
            ErrorList errors, MortgageParametersData parametersBusiness, bool calibrate)
        {
            object[,] values = null;
            if (File.Exists(fileName))
            {
                try
                {
                    values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadCurrentRegionValues(fileName, sheetName, "A1");
                }
                catch (Exception exc)
                {
                    throw new ApplicationException("Fout tijdens inlezen van mortgage data uit sheet " + sheetName + " in bestand " + fileName);
                }
            }
            DataCreator dataCreator = new DataCreator();
            PositionList positions = new PositionList();
            //            PositionList positions = dataCreator.ReadMortgagePositions_mortgagePackages(values, sheetName, errors, parametersBusiness, 
            //                calibrate);
            return positions;
        }
        public PositionList ReadMortgagePositions(DateTime reportDate, object[,] values, string sheetName,
            ErrorList errors, MortgageParametersData parametersBusiness, bool calibrate, bool bSkipFirstCoupon_MortgageModel, List<string> listPortfolios)
        {
            DataCreator dataCreator = new DataCreator();
            PositionList positions = dataCreator.ReadMortgagePositions(values, sheetName, errors, parametersBusiness, calibrate, bSkipFirstCoupon_MortgageModel, listPortfolios);
            return positions;
        }
        public MortgageParametersData ReadMortgageParameters(DateTime dtReport, string fileName_ModelParameters, string fileName_ActuariesData)
        {
            Dictionary<string, object[,]> MortgageParametersDataOrig = new Dictionary<string, object[,]>();
            if (File.Exists(fileName_ModelParameters))
            {
                ExcelWrapper.ExcelWrapper ExcelObj = new ExcelWrapper.ExcelWrapper();
                Workbook inputFile = null;
                Worksheet tab;
                Range range;
                try
                {
                    try
                    {
                        inputFile = ExcelObj.WorkbookOpen(fileName_ModelParameters, true);
                    }
                    catch
                    {
                        throw new IOException("Bestand '" + fileName_ModelParameters + "' kan niet worden geopend");
                    }

                    try
                    {
                        tab = inputFile.Worksheets["Input disconteringspread"];
                        range = ExcelWrapper.ExcelWrapper.RangeSetStartingAt(tab, "A1");
                        MortgageParametersDataOrig["Input disconteringspread"] = ExcelWrapper.ExcelWrapper.ReadRange(range);
                    }
                    catch
                    {
                        throw new IOException("Sheet 'Input disconteringspread' in bestand '" + fileName_ModelParameters + "' kan niet worden geopend");
                    }

                    try
                    {
                        tab = inputFile.Worksheets["Input Tarief"];
                        range = ExcelWrapper.ExcelWrapper.RangeSetStartingAt(tab, "A1");
                        MortgageParametersDataOrig["Input Tarief"] = ExcelWrapper.ExcelWrapper.ReadRange(range);
                    }
                    catch
                    {
                        throw new IOException("Sheet 'Input Tarief' in bestand '" + fileName_ModelParameters + "' kan niet worden geopend");
                    }
                    try
                    {
                        tab = inputFile.Worksheets["Input verval"];
                        range = ExcelWrapper.ExcelWrapper.RangeSetStartingAt(tab, "A1");
                        MortgageParametersDataOrig["Input verval"] = ExcelWrapper.ExcelWrapper.ReadRange(range);
                    }
                    catch
                    {
                        throw new IOException("Sheet 'Input verval' in bestand '" + fileName_ModelParameters + "' kan niet worden geopend");
                    }

                    try
                    {
                        tab = inputFile.Worksheets["Input other"];
                        range = ExcelWrapper.ExcelWrapper.RangeSetStartingAt(tab, "A1");
                        MortgageParametersDataOrig["Input other"] = ExcelWrapper.ExcelWrapper.ReadRange(range);
                    }
                    catch
                    {
                        throw new IOException("Sheet 'Input other' in bestand '" + fileName_ModelParameters + "' kan niet worden geopend");
                    }

                }
                finally
                {
                    inputFile.Close();
                    ExcelObj.Dispose();
                }

            }
            Dictionary<string, object[,]> ActuariesData = new Dictionary<string, object[,]>();
            if (File.Exists(fileName_ActuariesData))
            {
                ExcelWrapper.ExcelWrapper ExcelObj = new ExcelWrapper.ExcelWrapper();
                Workbook inputFile = null;
                Worksheet tab;
                Range range;
                try
                {
                    try
                    {
                        inputFile = ExcelObj.WorkbookOpen(fileName_ActuariesData, true);
                    }
                    catch
                    {
                        throw new IOException("Bestand '" + fileName_ActuariesData + "' kan niet worden geopend");
                    }

                    try
                    {
                        tab = inputFile.Worksheets["AG 2018 qx mannen"];
                        range = ExcelWrapper.ExcelWrapper.RangeSetStartingAt(tab, "A3");
                        ActuariesData["qx man"] = ExcelWrapper.ExcelWrapper.ReadRange(range);
                    }
                    catch
                    {
                        throw new IOException("Sheet 'AG 2018 qx mannen' in bestand '" + fileName_ActuariesData + "' kan niet worden geopend");
                    }

                    try
                    {
                        tab = inputFile.Worksheets["AG 2018 qx vrouwen"];
                        range = ExcelWrapper.ExcelWrapper.RangeSetStartingAt(tab, "A3");
                        ActuariesData["qx woman"] = ExcelWrapper.ExcelWrapper.ReadRange(range);
                    }
                    catch
                    {
                        throw new IOException("Sheet 'AG 2018 qx vrouwen' in bestand '" + fileName_ActuariesData + "' kan niet worden geopend");
                    }
                    try
                    {
                        tab = inputFile.Worksheets["CVS-ervaringsfactoren Hyp"];
                        range = ExcelWrapper.ExcelWrapper.RangeSetStartingAt(tab, "A3");
                        ActuariesData["ervaringsfactoren"] = ExcelWrapper.ExcelWrapper.ReadRange(range);
                    }
                    catch
                    {
                        throw new IOException("Sheet 'CVS-ervaringsfactoren Hyp' in bestand '" + fileName_ActuariesData + "' kan niet worden geopend");
                    }

                }
                finally
                {
                    inputFile.Close();
                    ExcelObj.Dispose();
                }

            }
            TotalRisk.MortgageModel.DataCreator dataCreator = new DataCreator();
            MortgageParametersData parameters = null;
            try
            {
                parameters = dataCreator.ReadMortgageParameters(dtReport, MortgageParametersDataOrig, ActuariesData);
            }
            catch
            {
                throw new IOException("Mortgage parameters data are no good");
            }
            return parameters;
        }
        // Utilities:
        public double GetMacDuration(double Coupon, double Yield, DateTime? maturityDate, DateTime? reportDate)
        {
            double ReturnValue;
            System.TimeSpan T = new System.TimeSpan();
            T = ((maturityDate.Value - reportDate.Value));
            double Years = ((double)T.Days) / 365.0;
            ReturnValue = Years;
            if (Yield > 0)
            {
                ReturnValue = (1 + Yield) / Yield;
                ReturnValue -= (((100 * (1 + Yield)) + Years * (Coupon - (100 * Yield))) /
                 (Coupon * (Math.Pow(1 + Yield, Years) - 1) + (100 * Yield)));
            }
            /*  
             *        1+Y       100(1+Y)+Years(C-100Y)
             *  D =  -----  -   -----------------------
             *         Y        C((1+Y)^T -1)+100Y
             * */

            return ReturnValue;
        }
        // DNB Stress test shocks for Bonds:
        public Tuple<double[], Dictionary<string, double[]>> Read_DNB_Stress_Shocks_GovBonds(object[,] values)
        {
            Dictionary<string, int> headers = HeaderNamesColumns(values); // import headers
            int nMaturities = headers.Count() - 2;
            double[] fMaturities = new double[nMaturities];
            int firstColumnValues = headers["Country Code".ToUpper()] + 2;
            for (int col = firstColumnValues; col <= values.GetUpperBound(DimensionCol); col++)
            {
                string mat = ReadFieldAsString(values, 1, col).Trim();
                mat = mat.Replace("Y", "");
                try
                {
                    double d = 0;
                    if (mat != "")
                    {
                        double.TryParse(mat, out d);
                    }
                    else
                    {
                        throw new ApplicationException("Error converting value to double in row = " + 1 + ", column " + col + " : " + mat);
                    }
                    fMaturities[col - firstColumnValues] = d;
                }
                catch (Exception exc)
                {
                    throw new ApplicationException("Error converting value to double in row = " + 1 + ", column " + col + " : " + exc.Message);
                }

            }
            Dictionary<string, double[]> table = new Dictionary<string, double[]>();
            for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
            {
                string countryCode = ReadFieldAsString(values, row, 0, headers, "Country Code").Trim().ToUpper();

                if (!table.ContainsKey(countryCode))
                {
                    double[] shocks = new double[nMaturities];
                    for (int col = firstColumnValues; col <= values.GetUpperBound(DimensionCol); col++)
                    {
                        shocks[col - firstColumnValues] = ReadFieldAsDouble(values, row, col) / 10000; // given in Bps
                    }
                    table.Add(countryCode, shocks);
                }
                else
                {
                    throw new ApplicationException("Error: second line with the same Country Code in row = " + row);
                }
            }
            return Tuple.Create(fMaturities, table);
        }
        public Tuple<Dictionary<string, Dictionary<string, double>>, Dictionary<string, Dictionary<string, double>>> Read_DNB_Stress_Shocks_CorBonds(object[,] values)
        {
            Dictionary<string, int> headers = HeaderNamesColumns(values); // import headers
            int nMaturities = headers.Count() - 2;
            double[] fMaturities = new double[nMaturities];
            int firstColumnValues = headers["Country Code".ToUpper()] + 2;
            Dictionary<string, Dictionary<string, double>> table_financial = new Dictionary<string, Dictionary<string, double>>();
            Dictionary<string, Dictionary<string, double>> table_nonFinancial = new Dictionary<string, Dictionary<string, double>>();
            for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
            {
                string countryCode = ReadFieldAsString(values, row, 0, headers, "Country Code").Trim().ToUpper();
                string type = ReadFieldAsString(values, row, 0, headers, "Type").Trim().ToUpper();
                Dictionary<string, Dictionary<string, double>> table;
                if (type == "FINANCIAL")
                {
                    table = table_financial;
                }
                else
                {
                    table = table_nonFinancial;
                }
                if (!table.ContainsKey(countryCode))
                {
                    table[countryCode] = new Dictionary<string, double>();
                    for (int col = firstColumnValues; col <= values.GetUpperBound(DimensionCol); col++)
                    {
                        string rating = ReadFieldAsString(values, 1, col).Trim().ToUpper();
                        table[countryCode][rating] = ReadFieldAsDouble(values, row, col) / 10000; // given in Bps
                    }
                }
                else
                {
                    throw new ApplicationException("Error: second line with the same Country Code in row = " + row);
                }
            }
            return Tuple.Create(table_financial, table_nonFinancial);
        }
        public Tuple<Dictionary<string, Dictionary<string, double>>> Read_DNB_Stress_Shocks_CovBonds(object[,] values)
        {
            Dictionary<string, int> headers = HeaderNamesColumns(values); // import headers
            int nMaturities = headers.Count() - 2;
            double[] fMaturities = new double[nMaturities];
            int firstColumnValues = headers["Country Code".ToUpper()] + 1;
            Dictionary<string, Dictionary<string, double>> table = new Dictionary<string, Dictionary<string, double>>();
            for (int row = 2; row <= values.GetUpperBound(DimensionRow); row++)
            {
                string countryCode = ReadFieldAsString(values, row, 0, headers, "Country Code").Trim().ToUpper();
                if (!table.ContainsKey(countryCode))
                {
                    table[countryCode] = new Dictionary<string, double>();
                    for (int col = firstColumnValues; col <= values.GetUpperBound(DimensionCol); col++)
                    {
                        string rating = ReadFieldAsString(values, 1, col).Trim().ToUpper();
                        table[countryCode][rating] = ReadFieldAsDouble(values, row, col) / 10000; // given in Bps
                    }
                }
                else
                {
                    throw new ApplicationException("Error: second line with the same Country Code in row = " + row);
                }
            }
            return Tuple.Create(table);
        }
        public Dictionary<string, object[,]> ReadFile_To_Dictionary(string fileName, bool usedRange, string firstCell = "A1")
        {
            Dictionary<string, object[,]> data = new Dictionary<string, object[,]>();
            TotalRisk.ExcelWrapper.ExcelWrapper ExcelObj = new TotalRisk.ExcelWrapper.ExcelWrapper();
            Workbook inputFile;
            Worksheet tab;
            Range range;
            if (File.Exists(fileName))
            {
                try
                {
                    try
                    {
                        inputFile = ExcelObj.WorkbookOpen(fileName, true);
                    }
                    catch
                    {
                        throw new IOException("Bestand '" + fileName + "' kan niet worden geopend");
                    }
                    if (null != inputFile.Worksheets)
                    {
                        for (int i = 1; i <= inputFile.Worksheets.Count; i++)
                        {
                            tab = inputFile.Worksheets[i];
                            if (usedRange)
                            {
                                range = tab.UsedRange;
                            }
                            else
                            {
                                range = TotalRisk.ExcelWrapper.ExcelWrapper.RangeSetCurrentRegion(tab, firstCell);
                            }
                            object[,] values = TotalRisk.ExcelWrapper.ExcelWrapper.ReadRange(tab, range);
                            data.Add(tab.Name, values);
                        }
                    }
                    else
                    {
                        inputFile.Close();
                        throw new IOException("File '" + fileName + "' has no worksheets");
                    }
                    inputFile.Close();
                }
                finally
                {
                    ExcelObj.Dispose();
                }
            }
            return data;
        }

    }
}
