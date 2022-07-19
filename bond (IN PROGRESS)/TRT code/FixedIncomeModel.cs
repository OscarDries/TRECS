using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using TotalRisk.ExcelWrapper;
using TotalRisk.Utilities;


namespace TotalRisk.ValuationModule
{
    public class FixedIncomeModels
    {
        public static string m_Currency = "EUR";
        public static Boolean m_HullWhitModel = false;
        public DateTime[] m_dStandardizedCashFlow_DatePoints;
        public SortedList<string, double> m_CurrencyHedgePercentage_Fund;
        public DNB_StressShocks m_StressShocks_DNB = null;
        // Scope || Portfolio ID | Security ID | Security ID LL | CBondDebugDataList
        public Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, CBondDebugDataList>>>> m_Scope_PortID_SecID_SecIDLL_BondDebugList_A = null;
        public FixedIncomeModels(SortedList<string, double> CurrencyHedgePercentage_Fund)
        {
            m_CurrencyHedgePercentage_Fund = CurrencyHedgePercentage_Fund;
        }

        public PositionList ProcessWithImport(FormMain frm, string name, DateTime dtReport, ScenarioList scenarios,
            Dictionary<string, CurveList> discountCurves, string fileNamePositions, string fileNameCashFlows, bool isCashFlowNeeded)
        {
            PositionList positions = new PositionList();
            ErrorList errors = new ErrorList();
            if (fileNamePositions != "")
            {
                m_dStandardizedCashFlow_DatePoints = getReportedCashFlowDates(dtReport, "Y", 50);
                positions = ImportPositions(frm, name, dtReport, scenarios, discountCurves, errors, fileNamePositions, fileNameCashFlows);
                Process(frm, name, dtReport, positions, scenarios, discountCurves, errors, isCashFlowNeeded);
            }

            return positions;
        }

        public PositionList ImportPositions(FormMain frm, string name, DateTime dtReport, ScenarioList scenarios,
            Dictionary<string, CurveList> discountCurves, ErrorList errors, string fileNamePositions, string fileNameCashFlows)
        {
            PositionList positions = new PositionList();

            if (fileNamePositions != "")
            {
                // Read positions
                frm.SetStatus("Inlezen posities " + name);
                Import_ValuationModule import = new Import_ValuationModule();
                if (name.Equals("Bond Forwards") && "" != fileNameCashFlows)
                {
                    positions = import.ReadBondForwardPositions_IMW(dtReport, fileNamePositions, fileNameCashFlows, scenarios, errors);
                }
                else if (name.Equals("Bonds") && "" != fileNameCashFlows)
                {
                    positions = import.ReadBondPositions_IMW(dtReport, frm.m_sReportingPeriod, fileNamePositions, fileNameCashFlows, scenarios, this, errors);
                }
                else if (name.Equals("Cash"))
                {
                    positions = import.ReadCashPositions_GARC(dtReport, fileNamePositions, errors);
                }
                else if (name.Equals("Swaps"))
                {
                    positions = import.ReadSwapPositions_IMW(dtReport, fileNamePositions, fileNameCashFlows, scenarios, errors);
                }
                else if (name.Equals("Swaptions"))
                {
                    positions = import.ReadSwaptionPositions_IMW(dtReport, fileNamePositions, scenarios, m_Currency, m_HullWhitModel, errors);
                }
                else if (name.Equals("Ad-Hoc Swaps"))
                {
                    positions = import.Read_Ad_Hoc_Positions_Assets(dtReport, fileNamePositions, scenarios, "SWAP", m_HullWhitModel, errors);
                }
                else if (name.Equals("Ad-Hoc Swaptions"))
                {
                    positions = import.Read_Ad_Hoc_Positions_Assets(dtReport, fileNamePositions, scenarios, "SWAPTION", m_HullWhitModel, errors);
                }
                else if (name.Equals("Spaarlossen"))
                {
                    frm.SetStatus("Inlezen Spaarlossen posities " + name);
                    positions = import.Read_Spaarlos_Positions(dtReport, fileNamePositions, scenarios, discountCurves, false, errors);
                }
                else if (name.Equals("Spaarlossen RiskMargin"))
                {
                    frm.SetStatus("Inlezen Spaarlossen posities " + name);
                    positions = import.Read_Spaarlos_Positions(dtReport, fileNamePositions, scenarios, discountCurves, true, errors);
                }
                else
                {
                    positions = import.ReadFixedIncomePositions(dtReport, fileNamePositions, scenarios, errors);
                }
            }

            return positions;
        }

        public void Process(FormMain frm, string name, DateTime dtReport, PositionList positions,
            ScenarioList scenarios, Dictionary<string, CurveList> discountCurves, ErrorList errors, bool isCashFlowNeeded)
        {
            if (name.Equals("Bond Forwards"))
            {
                Process_BondForward(frm, name, dtReport, positions, scenarios, errors, isCashFlowNeeded);
            }
            else if (name.Equals("Bonds"))
            {
                Process_Bond(frm, name, dtReport, positions, scenarios, errors, isCashFlowNeeded);
            }
            else if (name.Equals("Cash"))
            {
                Process_Cash(frm, name, dtReport, positions, scenarios, errors, isCashFlowNeeded);
            }
            else if (name.Equals("Swaps"))
            {
                Process_Swap(frm, name, dtReport, positions, scenarios, errors, isCashFlowNeeded);
            }
            else if (name.Equals("Ad-Hoc Swaps"))
            {
                Process_Swap(frm, name, dtReport, positions, scenarios, errors, isCashFlowNeeded);
            }
            else if (name.Equals("Swaptions"))
            {
                Process_Swaptions(frm, name, dtReport, positions, scenarios, errors);
            }
            else if (name.Equals("Spaarlossen"))
            {
                Process_Spaarlossen(frm, name, dtReport, positions, scenarios, discountCurves, errors, isCashFlowNeeded);
            }
            else if (name.Equals("Ad-Hoc Swaptions"))
            {
                Process_Swaptions(frm, name, dtReport, positions, scenarios, errors);
            }
            else
            {
                Process_FI(frm, name, dtReport, positions, scenarios, errors, isCashFlowNeeded);
            }
        }

        public void Process_Spaarlossen_RiskMargin(FormMain frm, string name, DateTime dtReport, PositionList positions,
            ScenarioList scenarios, Dictionary<string, CurveList> discountCurves, ErrorList errors, bool isCashFlowNeeded)
        {
            int cashflowCount = 0;
            int cashFlowCount_Original = 0;
            object[,] expectedCashflowValues = null;
            object[,] expectedCashflowValues_Original = null;
            try
            {
                object[,] scenarioValues = new object[positions.Count + 1, scenarios.Count];
                object[,] debugValues = new object[positions.Count + 1, 29];
                if (isCashFlowNeeded)
                {
                    cashflowCount = m_dStandardizedCashFlow_DatePoints.Length;
                    expectedCashflowValues = new object[positions.Count + 1, cashflowCount + 1];
                    expectedCashflowValues[0, 0] = "MarketValue";
                    for (int idxCashflow = 0; idxCashflow < cashflowCount; idxCashflow++)
                    {
                        DateTime dateHeader = m_dStandardizedCashFlow_DatePoints[idxCashflow];
                        expectedCashflowValues[0, idxCashflow + 1] = dateHeader;
                    }
                    Instrument_Cashflow instrument;
                    int indexSec = -1;
                    for (int idxPosition = 0; idxPosition < positions.Count; idxPosition++)
                    {
                        Position position = positions[idxPosition];
                        instrument = (Instrument_Cashflow)position.m_Instrument;
                        if (instrument.m_CashFlowSchedule.Count() > cashFlowCount_Original)
                        {
                            cashFlowCount_Original = instrument.m_CashFlowSchedule.Count();
                            indexSec = idxPosition;
                        }
                    }
                    expectedCashflowValues_Original = new object[positions.Count + 1, cashFlowCount_Original + 1];
                    expectedCashflowValues_Original[0, 0] = "MarketValue";
                    SortedList<DateTime, Cashflow> cashFlows = ((Instrument_Cashflow)positions[indexSec].m_Instrument).m_CashFlowSchedule.m_cashflows;
                    for (int idxCashflow = 0; idxCashflow < cashFlows.Count(); idxCashflow++)
                    {
                        DateTime dateHeader = cashFlows.Values[idxCashflow].m_Date;
                        expectedCashflowValues_Original[0, idxCashflow + 1] = dateHeader;
                    }

                }

                DateTime startTime = DateTime.Now;
                if (errors.CountErrors() == 0)
                {
                    int colDebug = 0;
                    debugValues[0, colDebug++] = "Period";
                    debugValues[0, colDebug++] = "DATA SOURCE";
                    debugValues[0, colDebug++] = "Row ID";
                    debugValues[0, colDebug++] = "Actuaries Scenario Name";
                    debugValues[0, colDebug++] = "Actuaries Scenario ID";
                    debugValues[0, colDebug++] = "CIC_LL";
                    debugValues[0, colDebug++] = "PositionId";
                    debugValues[0, colDebug++] = "PortfolioId";
                    debugValues[0, colDebug++] = "SecurityName";
                    debugValues[0, colDebug++] = "Spaar/Hybride";
                    debugValues[0, colDebug++] = "InstrumentType";
                    debugValues[0, colDebug++] = "CouponType";
                    debugValues[0, colDebug++] = "MaturityDate";
                    debugValues[0, colDebug++] = "Currency";
                    debugValues[0, colDebug++] = "Nominal";
                    debugValues[0, colDebug++] = "MarketValue";
                    debugValues[0, colDebug++] = "ImpliedSpread";
                    debugValues[0, colDebug++] = "Iterations";
                    debugValues[0, colDebug++] = "Duration";
                    debugValues[0, colDebug++] = "Value with Zero Spread";
                    debugValues[0, colDebug++] = "CollateralType";
                    debugValues[0, colDebug++] = "CollateralCoveragePerc";
                    debugValues[0, colDebug++] = "Scope3";
                    debugValues[0, colDebug++] = "Discount Curve ID";
                    debugValues[0, colDebug++] = "Tegenpartij";
                    debugValues[0, colDebug++] = "Hoofdpartij";
                    debugValues[0, colDebug++] = "Counterparty Name";
                    debugValues[0, colDebug++] = "Counterparty LEI";
                    debugValues[0, colDebug++] = "CQS";
                    debugValues[0, colDebug++] = "SmS";
                    debugValues[0, colDebug++] = "Message";

                    // Initialize progress bar
                    frm.InitProgBar(scenarios.Count);

                    Curve scenarioZeroCurve;
                    string ccy;
                    Scenario scenario;
                    double extraCreditSpread_SCRCharged;
                    double extraCreditSpread_Governmanets;
                    // Calculate market value for all combinations of positions/scenarios
                    for (int idxScenario = 0; idxScenario < scenarios.Count; idxScenario++)
                    {
                        scenario = scenarios[idxScenario];
                        scenarioValues[0, idxScenario] = scenario.m_sName;
                        frm.SetStatus("Verwerking " + name + " scenario " + scenario.m_sName);
                        CurveList scenarioZeroCurves = new CurveList();
                        foreach (ScenarioCurve scenarioCurve in scenario.m_YieldCurves)
                        {
                            ccy = scenarioCurve.m_sName;
                            scenarioZeroCurve = scenarioCurve.m_Curve;
                            scenarioZeroCurves.Add(ccy.ToUpper(), scenarioZeroCurve);
                        }
                        extraCreditSpread_SCRCharged = scenario.GetExtraCreditSpread("SCRCharged");
                        extraCreditSpread_Governmanets = scenario.GetExtraCreditSpread("Governments");
                        if (extraCreditSpread_Governmanets > 0)
                        {
                            extraCreditSpread_Governmanets += 0;
                        }
                        // Process all positions
                        for (int idxPosition = 0; idxPosition < positions.Count; idxPosition++)
                        {
                            Position position = positions[idxPosition];
                            Instrument_Cashflow instrument = (Instrument_Cashflow)position.m_Instrument;

                            // Create scenario discount curve
                            ccy = Position.TranslateCurrency_Curve(scenarioZeroCurves, position.m_sCurrency);
                            scenarioZeroCurve = scenarioZeroCurves[ccy];

                            // Calculate market value for this position with given scenario curve
                            double modelValueScenario = 0;
                            if (null != position.m_SpreadRiskData)
                            {
                                if (position.m_SpreadRiskData.m_bSCRStress)
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread);
                                    modelValueScenario *= (1 - position.m_SpreadRiskData.m_fSpreadDuration * extraCreditSpread_SCRCharged);
                                }
                                else if ("Bond" == position.m_SpreadRiskData.m_sSCRLevel1Type &&
                                    "Bond_Government" == position.m_SpreadRiskData.m_sASRLevel2Type)
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread + extraCreditSpread_Governmanets);
                                }
                                else
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread);
                                }
                            }
                            else
                            {
                                if ("" != position.m_sCIC_LL && "A2" == position.m_sCIC_LL.Substring(2, 2))
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread + extraCreditSpread_Governmanets);
                                }
                                else
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread);
                                }
                            }
                            double marketValueScenario = instrument.m_fDirtyValue + (modelValueScenario - instrument.m_fDirtyValue);
                            scenarioValues[idxPosition + 1, idxScenario] = modelValueScenario;
                            if (scenario.isFairValueScenario())
                            {
                                colDebug = 0;
                                debugValues[idxPosition + 1, colDebug++] = "'" + frm.m_sReportingPeriod;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sDATA_Source;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sRow;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sActuariesScenarioName;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sActuariesScenarioID;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCIC_LL;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sPositionId;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sPortfolioId;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSecurityName;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSecurityType;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_sInstrumentType;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_sCouponType;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_MaturityDate;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCurrency;
                                debugValues[idxPosition + 1, colDebug++] = position.m_fVolume;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fDirtyValue;
                                if (double.IsNaN(instrument.m_fImpliedSpread))
                                {
                                    debugValues[idxPosition + 1, colDebug++] = "#N/A";
                                }
                                else
                                {
                                    debugValues[idxPosition + 1, colDebug++] = instrument.m_fImpliedSpread;
                                }
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_dIterations;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fDuration;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fValueAtZeroSpread;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCollateralType;
                                debugValues[idxPosition + 1, colDebug++] = position.m_fCollateralCoveragePercentage;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sScope3;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_sDiscountCurveCode;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCounterpartyIssuer_Name;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCounterpartyGroup_Name;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCounterparty_Name;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCounterparty_LEI;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sIssuerCreditQuality;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSMS_Entity;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sMessage;
                                position.m_fFairValue = modelValueScenario;

                                if (isCashFlowNeeded)
                                {
                                    instrument.calculateStandardizedCashFlows(dtReport, m_dStandardizedCashFlow_DatePoints);
                                    expectedCashflowValues[idxPosition + 1, 0] = marketValueScenario;
                                    for (int idxCashflow = 0; idxCashflow < cashflowCount; idxCashflow++)
                                    {
                                        expectedCashflowValues[idxPosition + 1, idxCashflow + 1] = instrument.m_CashFlow_Reported[idxCashflow];
                                    }
                                    expectedCashflowValues_Original[idxPosition + 1, 0] = marketValueScenario;
                                    SortedList<DateTime, Cashflow> cashFlows = ((Instrument_Cashflow)position.m_Instrument).m_CashFlowSchedule.m_cashflows;
                                    for (int idxCashflow = 0; idxCashflow < cashFlows.Count(); idxCashflow++)
                                    {
                                        expectedCashflowValues_Original[idxPosition + 1, idxCashflow + 1] = cashFlows.Values[idxCashflow].m_fAmount;
                                    }
                                }
                            }

                            System.Windows.Forms.Application.DoEvents();
                        }

                        // Update progress bar
                        frm.IncrementProgBar();
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                DateTime endTime = DateTime.Now;
                double runTime = (endTime - startTime).TotalSeconds;

                // Write aggregated market values to results workbook
                frm.SetStatus("Opslaan resultaten " + name);
                Scenario baseScenario = scenarios.getScenarioFairValue();
                CurveList baseCurves = new CurveList();
                foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
                {
                    baseCurves.Add(scenarioCurve.m_sName.ToUpper(), scenarioCurve.m_Curve);
                }
                foreach (KeyValuePair<string, Curve> dCurve in discountCurves["EUR"])
                {
                    baseCurves.Add("DiscountCurve: " + dCurve.Key.ToUpper(), dCurve.Value);
                }
                object[,] curveValues = baseCurves.ToArray(true);
                object[,] errorValues = errors.ToArray();
                frm.WriteToExcel(name, positions, scenarioValues, curveValues, debugValues, errorValues, runTime, null, null, false);
                if (isCashFlowNeeded)
                {
                    frm.WriteToExcel(name + " Cash flows (EUR)", positions, expectedCashflowValues, curveValues, debugValues, errorValues, runTime, null, null, false);
                    frm.WriteToExcel(name + " Cash flows (EUR) Original", positions, expectedCashflowValues_Original, curveValues, debugValues, errorValues, runTime, null, null, false);
                }

                if (errors.CountErrors() > 0)
                {
                    MessageBox.Show("Fouten in verwerking " + name + ", bekijk Error werkblad in output bestand", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (IOException exc)
            {
                MessageBox.Show("Fout in verwerking " + name + ":\n" + exc.Message, "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            frm.HideProgBar();
        }
        public void Process_Spaarlossen(FormMain frm, string name, DateTime dtReport, PositionList positions,
            ScenarioList scenarios, Dictionary<string, CurveList> discountCurves, ErrorList errors, bool isCashFlowNeeded)
        {
            int cashflowCount = 0;
            int cashFlowCount_Original = 0;
            object[,] expectedCashflowValues = null;
            object[,] expectedCashflowValues_Original = null;
            try
            {
                List<string> headerNames = new List<string>();
                headerNames.Add("Period");
                headerNames.Add("DATA SOURCE");
                headerNames.Add("Row ID");
                headerNames.Add("Actuaries Scenario Name");
                headerNames.Add("Actuaries Scenario ID");
                headerNames.Add("CIC_LL");
                headerNames.Add("SelectieIndex_LL_SCR");
                headerNames.Add("PositionId");
                headerNames.Add("PortfolioId");
                headerNames.Add("SecurityName");
                headerNames.Add("Spaar/Hybride");
                headerNames.Add("InstrumentType");
                headerNames.Add("CouponType");
                headerNames.Add("MaturityDate");
                headerNames.Add("Currency");
                headerNames.Add("Nominal");
                headerNames.Add("MarketValue");
                headerNames.Add("ImpliedSpread");
                headerNames.Add("Iterations");
                headerNames.Add("Duration");
                headerNames.Add("Value with Zero Spread");
                headerNames.Add("CollateralType");
                headerNames.Add("CollateralCoveragePerc");
                headerNames.Add("Scope3");
                headerNames.Add("Scope3_Issuer");
                headerNames.Add("Scope3_Investor");
                headerNames.Add("Discount Curve ID");
                headerNames.Add("Tegenpartij");
                headerNames.Add("Hoofdpartij");
                headerNames.Add("Counterparty Name");
                headerNames.Add("Counterparty LEI");
                headerNames.Add("CQS");
                headerNames.Add("SmS");
                headerNames.Add("Message");

                object[,] scenarioValues = new object[positions.Count + 1, scenarios.Count];
                object[,] debugValues = new object[positions.Count + 1, headerNames.Count];
                if (isCashFlowNeeded)
                {
                    cashflowCount = m_dStandardizedCashFlow_DatePoints.Length;
                    expectedCashflowValues = new object[positions.Count + 1, cashflowCount + 1];
                    expectedCashflowValues[0, 0] = "MarketValue";
                    for (int idxCashflow = 0; idxCashflow < cashflowCount; idxCashflow++)
                    {
                        DateTime dateHeader = m_dStandardizedCashFlow_DatePoints[idxCashflow];
                        expectedCashflowValues[0, idxCashflow + 1] = dateHeader;
                    }
                    Instrument_Cashflow instrument;
                    int indexSec = -1;
                    for (int idxPosition = 0; idxPosition < positions.Count; idxPosition++)
                    {
                        Position position = positions[idxPosition];
                        instrument = (Instrument_Cashflow)position.m_Instrument;
                        if (instrument.m_CashFlowSchedule.Count() > cashFlowCount_Original)
                        {
                            cashFlowCount_Original = instrument.m_CashFlowSchedule.Count();
                            indexSec = idxPosition;
                        }
                    }
                    expectedCashflowValues_Original = new object[positions.Count + 1, cashFlowCount_Original + 1];
                    expectedCashflowValues_Original[0, 0] = "MarketValue";
                    SortedList<DateTime, Cashflow> cashFlows = ((Instrument_Cashflow)positions[indexSec].m_Instrument).m_CashFlowSchedule.m_cashflows;
                    for (int idxCashflow = 0; idxCashflow < cashFlows.Count(); idxCashflow++)
                    {
                        DateTime dateHeader = cashFlows.Values[idxCashflow].m_Date;
                        expectedCashflowValues_Original[0, idxCashflow + 1] = dateHeader;
                    }

                }

                DateTime startTime = DateTime.Now;
                if (errors.CountErrors() == 0)
                {
                    int colDebug = 0;
                    for (colDebug = 0; colDebug < headerNames.Count; colDebug++)
                    {
                        debugValues[0, colDebug] = "'" + headerNames[colDebug];
                    }

                    // Initialize progress bar
                    frm.InitProgBar(scenarios.Count);

                    Curve scenarioZeroCurve;
                    string ccy;
                    Scenario scenario;
                    double extraCreditSpread_SCRCharged;
                    double extraCreditSpread_Governmanets;
                    // Calculate market value for all combinations of positions/scenarios
                    for (int idxScenario = 0; idxScenario < scenarios.Count; idxScenario++)
                    {
                        scenario = scenarios[idxScenario];
                        scenarioValues[0, idxScenario] = scenario.m_sName;
                        frm.SetStatus("Verwerking " + name + " scenario " + scenario.m_sName);
                        CurveList scenarioZeroCurves = new CurveList();
                        foreach (ScenarioCurve scenarioCurve in scenario.m_YieldCurves)
                        {
                            ccy = scenarioCurve.m_sName;
                            scenarioZeroCurve = scenarioCurve.m_Curve;
                            scenarioZeroCurves.Add(ccy.ToUpper(), scenarioZeroCurve);
                        }
                        extraCreditSpread_SCRCharged = scenario.GetExtraCreditSpread("SCRCharged");
                        extraCreditSpread_Governmanets = scenario.GetExtraCreditSpread("Governments");
                        if (extraCreditSpread_Governmanets > 0)
                        {
                            extraCreditSpread_Governmanets += 0;
                        }
                        // Process all positions
                        for (int idxPosition = 0; idxPosition < positions.Count; idxPosition++)
                        {
                            Position position = positions[idxPosition];
                            Instrument_Cashflow instrument = (Instrument_Cashflow)position.m_Instrument;

                            // Create scenario discount curve
                            ccy = Position.TranslateCurrency_Curve(scenarioZeroCurves, position.m_sCurrency);
                            scenarioZeroCurve = scenarioZeroCurves[ccy];

                            // Calculate market value for this position with given scenario curve
                            double modelValueScenario = 0;
                            if (null != position.m_SpreadRiskData)
                            {
                                if (position.m_SpreadRiskData.m_bSCRStress)
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread);
                                    modelValueScenario *= (1 - position.m_SpreadRiskData.m_fSpreadDuration * extraCreditSpread_SCRCharged);
                                }
                                else if ("Bond" == position.m_SpreadRiskData.m_sSCRLevel1Type &&
                                    "Bond_Government" == position.m_SpreadRiskData.m_sASRLevel2Type)
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread + extraCreditSpread_Governmanets);
                                }
                                else
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread);
                                }
                            }
                            else
                            {
                                if ("" != position.m_sCIC_LL && "A2" == position.m_sCIC_LL.Substring(2, 2))
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread + extraCreditSpread_Governmanets);
                                }
                                else
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread);
                                }
                            }
                            double marketValueScenario = instrument.m_fDirtyValue + (modelValueScenario - instrument.m_fDirtyValue);
                            scenarioValues[idxPosition + 1, idxScenario] = modelValueScenario;
                            if (scenario.isFairValueScenario())
                            {
                                colDebug = 0;
                                debugValues[idxPosition + 1, colDebug++] = "'" + frm.m_sReportingPeriod;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sDATA_Source;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sRow;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sActuariesScenarioName;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sActuariesScenarioID;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCIC_LL;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSelectieIndex_LL;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sPositionId;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sPortfolioId;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSecurityName;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSecurityType;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_sInstrumentType;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_sCouponType;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_MaturityDate;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCurrency;
                                debugValues[idxPosition + 1, colDebug++] = position.m_fVolume;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fDirtyValue;
                                if (double.IsNaN(instrument.m_fImpliedSpread))
                                {
                                    debugValues[idxPosition + 1, colDebug++] = "#N/A";
                                }
                                else
                                {
                                    debugValues[idxPosition + 1, colDebug++] = instrument.m_fImpliedSpread;
                                }
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_dIterations;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fDuration;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fValueAtZeroSpread;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCollateralType;
                                debugValues[idxPosition + 1, colDebug++] = position.m_fCollateralCoveragePercentage;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sScope3;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sScope3_Issuer;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sScope3_Investor;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_sDiscountCurveCode;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCounterpartyIssuer_Name;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCounterpartyGroup_Name;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCounterparty_Name;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCounterparty_LEI;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sIssuerCreditQuality;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSMS_Entity;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sMessage;
                                position.m_fFairValue = modelValueScenario;

                                if (isCashFlowNeeded)
                                {
                                    instrument.calculateStandardizedCashFlows(dtReport, m_dStandardizedCashFlow_DatePoints);
                                    expectedCashflowValues[idxPosition + 1, 0] = marketValueScenario;
                                    for (int idxCashflow = 0; idxCashflow < cashflowCount; idxCashflow++)
                                    {
                                        expectedCashflowValues[idxPosition + 1, idxCashflow + 1] = instrument.m_CashFlow_Reported[idxCashflow];
                                    }
                                    expectedCashflowValues_Original[idxPosition + 1, 0] = marketValueScenario;
                                    SortedList<DateTime, Cashflow> cashFlows = ((Instrument_Cashflow)position.m_Instrument).m_CashFlowSchedule.m_cashflows;
                                    for (int idxCashflow = 0; idxCashflow < cashFlows.Count(); idxCashflow++)
                                    {
                                        expectedCashflowValues_Original[idxPosition + 1, idxCashflow + 1] = cashFlows.Values[idxCashflow].m_fAmount;
                                    }
                                }
                            }

                            System.Windows.Forms.Application.DoEvents();
                        }

                        // Update progress bar
                        frm.IncrementProgBar();
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                DateTime endTime = DateTime.Now;
                double runTime = (endTime - startTime).TotalSeconds;

                // Write aggregated market values to results workbook
                frm.SetStatus("Opslaan resultaten " + name);
                Scenario baseScenario = scenarios.getScenarioFairValue();
                CurveList baseCurves = new CurveList();
                foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
                {
                    baseCurves.Add(scenarioCurve.m_sName.ToUpper(), scenarioCurve.m_Curve);
                }
                foreach (KeyValuePair<string, Curve> dCurve in discountCurves["EUR"])
                {
                    baseCurves.Add("DiscountCurve: " + dCurve.Key.ToUpper(), dCurve.Value);
                }
                object[,] curveValues = baseCurves.ToArray(true);
                object[,] errorValues = errors.ToArray();
                frm.WriteToExcel(name, positions, scenarioValues, curveValues, debugValues, errorValues, runTime, null, null, false);
                if (isCashFlowNeeded)
                {
                    frm.WriteToExcel(name + " Cash flows (EUR)", positions, expectedCashflowValues, curveValues, debugValues, errorValues, runTime, null, null, false);
                    frm.WriteToExcel(name + " Cash flows (EUR) Original", positions, expectedCashflowValues_Original, curveValues, debugValues, errorValues, runTime, null, null, false);
                }

                if (errors.CountErrors() > 0)
                {
                    MessageBox.Show("Fouten in verwerking " + name + ", bekijk Error werkblad in output bestand", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (IOException exc)
            {
                MessageBox.Show("Fout in verwerking " + name + ":\n" + exc.Message, "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            frm.HideProgBar();
        }
        public void Process_BondForward(FormMain frm, string name, DateTime dtReport, PositionList positions,
            ScenarioList scenarios, ErrorList errors, bool isCashFlowNeeded)
        {
            int cashflowCount = m_dStandardizedCashFlow_DatePoints.Length;
            try
            {
                object[,] expectedCashflowValues = new object[2 * positions.Count + 1, cashflowCount + 1];
                expectedCashflowValues[0, 0] = "MarketValue";
                for (int idxCashflow = 0; idxCashflow < cashflowCount; idxCashflow++)
                {
                    DateTime dateHeader = m_dStandardizedCashFlow_DatePoints[idxCashflow];
                    expectedCashflowValues[0, idxCashflow + 1] = dateHeader;
                }
                object[,] scenarioValues = new object[positions.Count + 1, scenarios.Count];
                object[,] debugValues = new object[positions.Count + 1, 34];

                DateTime startTime = DateTime.Now;
                if (errors.CountErrors() == 0)
                {
                    int colDebug = 0;
                    debugValues[0, colDebug++] = "PositionId";
                    debugValues[0, colDebug++] = "Scope3";
                    debugValues[0, colDebug++] = "PortfolioId";
                    debugValues[0, colDebug++] = "SecurityID_LL";
                    debugValues[0, colDebug++] = "SecurityName";
                    debugValues[0, colDebug++] = "InstrumentType";
                    debugValues[0, colDebug++] = "CIC LL";
                    debugValues[0, colDebug++] = "MaturityDate";
                    debugValues[0, colDebug++] = "StrikePrice";
                    debugValues[0, colDebug++] = "MarketPrice";
                    debugValues[0, colDebug++] = "Model Type";
                    debugValues[0, colDebug++] = "ImpliedRepoRate";
                    debugValues[0, colDebug++] = "Duration";
                    debugValues[0, colDebug++] = "Bond ID";
                    debugValues[0, colDebug++] = "Bond Name";
                    debugValues[0, colDebug++] = "DNB_Type";
                    debugValues[0, colDebug++] = "DNB_CountryUnionCode";
                    debugValues[0, colDebug++] = "DNB_SpreadShock";
                    debugValues[0, colDebug++] = "Bond In Default";
                    debugValues[0, colDebug++] = "Bond Matrket Price";
                    debugValues[0, colDebug++] = "Bond Model Price";
                    debugValues[0, colDebug++] = "Bond Implied Spread";
                    debugValues[0, colDebug++] = "Bond Implied Spread Found";
                    debugValues[0, colDebug++] = "Bond Maturity";
                    debugValues[0, colDebug++] = "Bond Maturity_in_Years";
                    debugValues[0, colDebug++] = "Bond Conversion Factor";
                    debugValues[0, colDebug++] = "Bond Coupon";
                    debugValues[0, colDebug++] = "Bond Coupon Freq";
                    debugValues[0, colDebug++] = "Bond Duration";
                    debugValues[0, colDebug++] = "CP Name";
                    debugValues[0, colDebug++] = "CP LEI";
                    debugValues[0, colDebug++] = "CP CQS";
                    debugValues[0, colDebug++] = "CollateralCoveragePerc";
                    debugValues[0, colDebug++] = "Message";

                    // Initialize progress bar
                    frm.InitProgBar(scenarios.Count);

                    Curve scenarioZeroCurve;
                    string ccy;
                    Scenario scenario;
                    double extraCreditSpread_SCRCharged;
                    double extraCreditSpread_Governmanets;
                    double fDNB_SpreadShock = 0;
                    // Calculate market value for all combinations of positions/scenarios
                    for (int idxScenario = 0; idxScenario < scenarios.Count; idxScenario++)
                    {
                        scenario = scenarios[idxScenario];
                        scenarioValues[0, idxScenario] = scenario.m_sName;
                        frm.SetStatus("Verwerking " + name + " scenario " + scenario.m_sName);
                        CurveList scenarioZeroCurves = new CurveList();
                        foreach (ScenarioCurve scenarioCurve in scenario.m_YieldCurves)
                        {
                            ccy = scenarioCurve.m_sName;
                            scenarioZeroCurve = scenarioCurve.m_Curve;
                            scenarioZeroCurves.Add(ccy.ToUpper(), scenarioZeroCurve);
                        }
                        extraCreditSpread_SCRCharged = scenario.GetExtraCreditSpread("SCRCharged");
                        extraCreditSpread_Governmanets = scenario.GetExtraCreditSpread("Governments");
                        if (extraCreditSpread_Governmanets > 0)
                        {
                            extraCreditSpread_Governmanets += 0;
                        }
                        // Process all positions
                        for (int idxPosition = 0; idxPosition < positions.Count; idxPosition++)
                        {
                            Position position = positions[idxPosition];
                            Instrument_BondForward instrument = (Instrument_BondForward)position.m_Instrument;
                            double fMaturity_UnderlyingBond = instrument.getTimeToDate(dtReport, instrument.m_MaturityDate_UnderlyingBond);
                            // Create scenario discount curve
                            ccy = Position.TranslateCurrency_Curve(scenarioZeroCurves, position.m_sCurrency);
                            scenarioZeroCurve = scenarioZeroCurves[ccy];
                            // DNB Spread shock:
                            double fImpliedSpread_UndelyingBond = instrument.m_fImpliedSpread_UndelyingBond;
                            fDNB_SpreadShock = 0;
                            // Check for DNB shock:
                            if (null != m_StressShocks_DNB)
                            {
                                if (instrument.m_OriginalBondForward.m_DNB_Type == DNB_Bond.Government)
                                {
                                    string sCountryCode = instrument.m_OriginalBondForward.m_sDNB_CountryUnion;
                                    if (m_StressShocks_DNB.m_GovermentBondsShocks.ContainsKey(sCountryCode))
                                    {
                                        fDNB_SpreadShock = getDNBShock(fMaturity_UnderlyingBond, m_StressShocks_DNB.m_fGovermentBondsMatirities, m_StressShocks_DNB.m_GovermentBondsShocks[sCountryCode]);
                                    }
                                    fImpliedSpread_UndelyingBond += fDNB_SpreadShock;
                                }
                            }
                            // Calculate market value for this position with given scenario curve
                            double modelValue = instrument.getPrice(dtReport, scenarioZeroCurve,
                            fImpliedSpread_UndelyingBond + extraCreditSpread_Governmanets, instrument.m_fFxRate);
                            double marketValueScenario = instrument.m_fMarketPrice + (modelValue - instrument.m_fModelPrice);
                            scenarioValues[idxPosition + 1, idxScenario] = marketValueScenario;
                            if (scenario.isFairValueScenario())
                            {
                                colDebug = 0;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sUniquePositionId;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sScope3;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sPortfolioId;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSecurityID_LL;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSecurityName_LL;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSecurityType_LL;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCIC_LL;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_MaturityDate;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fStrikePrice;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fMarketPrice;
                                debugValues[idxPosition + 1, colDebug++] = "BondForward Model " + instrument.m_dModetType;
                                if (double.IsNaN(instrument.m_fImpliedRepoSpread_UndelyingBond))
                                {
                                    debugValues[idxPosition + 1, colDebug++] = "#N/A";
                                }
                                else
                                {
                                    debugValues[idxPosition + 1, colDebug++] = instrument.m_fImpliedRepoSpread_UndelyingBond;
                                }
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fDuration;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_OriginalBondForward.m_sUnderlyingSecurityCode;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_OriginalBondForward.m_sUnderlyingSecurityName;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_OriginalBondForward.m_DNB_Type.ToString();
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_OriginalBondForward.m_sDNB_CountryUnion;
                                debugValues[idxPosition + 1, colDebug++] = fDNB_SpreadShock;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_bDefaulted_UnderlyingBond;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fMarketPrice_UnderlyingBond;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fModelPrice_UnderlyingBond;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fImpliedSpread_UndelyingBond;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_bImpliedSpreadFound_UnderlyingBond;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_MaturityDate_UnderlyingBond;
                                debugValues[idxPosition + 1, colDebug++] = fMaturity_UnderlyingBond;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fConversionFactor_UnderlyingBond;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fCouponPerc_UnderlyingBond;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fCouponFrequency_UnderlyingBond;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fDuration_UnderlyingBond;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_OriginalBondForward.m_sGroupCounterpartyName;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_OriginalBondForward.m_sGroupCounterpartyLEI;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_OriginalBondForward.m_sGroupCounterpartyCQS;
                                debugValues[idxPosition + 1, colDebug++] = position.m_fCollateralCoveragePercentage;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sMessage;

                                position.m_fFairValue = marketValueScenario;
                                if (isCashFlowNeeded)
                                {
                                    instrument.m_CashFlow_Reported = new double[2][];
                                    for (int i = 0; i < 2; i++)
                                    {
                                        instrument.calculateStandardizedCashFlows(dtReport, scenarioZeroCurve, m_dStandardizedCashFlow_DatePoints, i);
                                    }
                                    int row_1 = 2 * idxPosition + 1;
                                    int row_2 = 2 * idxPosition + 2;
                                    if (BondForwardType.Forward == instrument.m_Type)
                                    {
                                        expectedCashflowValues[row_1, 0] = (instrument.m_fStrikePrice + instrument.m_fQuotationPrice) * instrument.m_fNominal_UnderlyingBond * instrument.m_fFxRate; ; // underlying bond leg
                                        expectedCashflowValues[row_2, 0] = -instrument.m_fStrikePrice * instrument.m_fNominal_UnderlyingBond * instrument.m_fFxRate; ; // product payment leg
                                    }
                                    else
                                    {
                                        expectedCashflowValues[row_1, 0] = instrument.m_fQuotationPrice * instrument.m_fNominal_UnderlyingBond * instrument.m_fFxRate; ; // underlying bond leg
                                        expectedCashflowValues[row_2, 0] = -instrument.m_fQuotationPrice * instrument.m_fNominal_UnderlyingBond * instrument.m_fFxRate; ; // product payment leg
                                    }

                                    double[] CF_RiskRente_0 = instrument.m_CashFlow_Reported[0];
                                    double[] CF_RiskRente_1 = instrument.m_CashFlow_Reported[1];
                                    for (int idxCashflow = 0; idxCashflow < cashflowCount; idxCashflow++)
                                    {
                                        expectedCashflowValues[row_1, idxCashflow + 1] = CF_RiskRente_0[idxCashflow] * instrument.m_fNominal_UnderlyingBond * instrument.m_fFxRate;
                                        expectedCashflowValues[row_2, idxCashflow + 1] = CF_RiskRente_1[idxCashflow] * instrument.m_fNominal_UnderlyingBond * instrument.m_fFxRate;
                                    }
                                }
                            }

                            System.Windows.Forms.Application.DoEvents();
                        }

                        // Update progress bar
                        frm.IncrementProgBar();
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                DateTime endTime = DateTime.Now;
                double runTime = (endTime - startTime).TotalSeconds;

                // Write aggregated market values to results workbook
                frm.SetStatus("Opslaan resultaten " + name);
                Scenario baseScenario = scenarios.getScenarioFairValue();
                CurveList baseCurves = new CurveList();
                foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
                {
                    baseCurves.Add(scenarioCurve.m_sName.ToUpper(), scenarioCurve.m_Curve);
                }
                object[,] curveValues = baseCurves.ToArray(true);
                object[,] errorValues = errors.ToArray();
                frm.WriteToExcel(name, positions, scenarioValues, curveValues, debugValues, errorValues, runTime, null, null, false);
                if (isCashFlowNeeded)
                {
                    frm.WriteToExcel(name + " Cash flows (EUR)", positions, expectedCashflowValues, curveValues, debugValues, errorValues, runTime, null, null, true);
                }

                if (errors.CountErrors() > 0)
                {
                    MessageBox.Show("Fouten in verwerking " + name + ", bekijk Error werkblad in output bestand", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (IOException exc)
            {
                MessageBox.Show("Fout in verwerking " + name + ":\n" + exc.Message, "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            frm.HideProgBar();
        }
        public double getDNBShock(double mat, double[] maturities, double[] shocks)
        {// linear interpolation
            double value = 0;
            int n = maturities.Length;
            int index_L = 0;
            int index_H = 0;
            if (maturities[0] >= mat) return shocks[0];
            if (maturities[n-1] <= mat) return shocks[n-1];
            while (mat > maturities[index_H]) index_H++;
            index_L = index_H - 1;
            value = shocks[index_L] + (mat - maturities[index_L]) / (maturities[index_H] - maturities[index_L]) * (shocks[index_H]-shocks[index_L]);
            return value;
        }
        /** BONDS: */
        public CBondDebugDataList getDebugData_Bonds(object[,] values)
        {
            if (null == values) return null;
            Import_ValuationModule Import = new Import_ValuationModule();
            Dictionary<string, int> headers = Import.HeaderNamesColumns(values); // import headers
            CBondDebugDataList list = new CBondDebugDataList();

            for (int row = values.GetLowerBound(0) + 1; row <= values.GetUpperBound(0); row++)
            {
                CBondDebugData p = new CBondDebugData();
                p.m_sReportingPeriod = Import.ReadFieldAsString(values, row, 0, headers, "Period").Trim();
                if (Import.HeaderNameExists(headers, "Report Date"))
                {
                    p.m_ReportDate = Import.ReadFieldAsDateTime(values, row, 0, headers, "Report Date"); // in the IMW file
                }
                if (Import.HeaderNameExists(headers, "Row ID Source"))
                {
                    p.m_sRow_Source = Import.ReadFieldAsString(values, row, 0, headers, "Row ID Source").Trim(); // row id in the IMW file
                }
                else p.m_sRow_Source = "";
                if (Import.HeaderNameExists(headers, "Row ID Debug"))
                {
                    p.m_sRow_Debug = Import.ReadFieldAsString(values, row, 0, headers, "Row ID Debug").Trim(); ; // row id in the debug tab
                }
                else p.m_sRow_Debug = "";
                if (Import.HeaderNameExists(headers, "Row ID Debug Linked"))
                {
                    p.m_sRow_Debug_Linked = Import.ReadFieldAsString(values, row, 0, headers, "Row ID Debug Linked").Trim(); // row id in the debug tab of linked position of previous period
                }
                else p.m_sRow_Debug_Linked = "";
                p.m_dNumberOfLinks = 0;
                p.m_sScope3 = Import.ReadFieldAsString(values, row, 0, headers, "Scope3").Trim();
                p.m_bIsLookThroughData = Import.ReadFieldAsBool(values, row, 0, headers, "Is Look-Through Data");
                p.m_sPosition_ID = Import.ReadFieldAsString(values, row, 0, headers, "PositionId").Trim();
                p.m_sPortfolio_ID = Import.ReadFieldAsString(values, row, 0, headers, "PortfolioId").Trim();
                p.m_sSecurity_ID = Import.ReadFieldAsString(values, row, 0, headers, "SecurityID").Trim();
                p.m_sSecurity_ID_LL = Import.ReadFieldAsString(values, row, 0, headers, "SecurityID_LL").Trim();
                p.m_sSecurity_Type = Import.ReadFieldAsString(values, row, 0, headers, "InstrumentType").Trim();
                p.m_sCIC = Import.ReadFieldAsString(values, row, 0, headers, "CIC").Trim();
                p.m_sCIC_LL = Import.ReadFieldAsString(values, row, 0, headers, "CIC_LL").Trim();

                p.m_fNominal_Value = Import.ReadFieldAsDouble(values, row, 0, headers, "Nominal_Value");
                p.m_fMarket_Value = Import.ReadFieldAsDouble(values, row, 0, headers, "Market_Value");
                p.m_fImplied_Spread = Import.ReadFieldAsDouble(values, row, 0, headers, "Implied_Spread");
                p.m_fFxRate = Import.ReadFieldAsDouble(values, row, 0, headers, "FX Rate");
                list.Add(p);
            }
            return list;
        }
        public CBondDebugData getDebugData_Bond(string period, Position position)
        {
            CBondDebugData p = new CBondDebugData();
            Instrument_Bond instrument = (Instrument_Bond)position.m_Instrument;
            Instrument_Bond_OriginalData m_OriginalBond = instrument.m_OriginalBond;
            p.m_sReportingPeriod = period;
            p.m_ReportDate = m_OriginalBond.m_dtReport;
            p.m_sRow_Source = position.m_sRow; // row id in the IMW file
            p.m_sRow_Debug = ""; // row id in the debug tab
            p.m_sRow_Debug_Linked = ""; // row id in the debug tab of linked position of previous period
            p.m_dNumberOfLinks = 0; // the number of lincked positions
            p.m_sScope3 = position.m_sScope3;
            p.m_bIsLookThroughData = position.m_bIsLookThroughPosition;
            p.m_sPosition_ID = position.m_sUniquePositionId;
            p.m_sPortfolio_ID = position.m_sPortfolioId;
            p.m_sSecurity_ID = position.m_sSecurityID;
            p.m_sSecurity_ID_LL = position.m_sSecurityID_LL;
            p.m_sSecurity_Type = position.m_sSecurityType_LL;
            p.m_sCIC = position.m_sCIC;
            p.m_sCIC_LL = position.m_sCIC_LL;

            p.m_fNominal_Value = instrument.m_fNominal;
            p.m_fMarket_Value = instrument.m_fMarketPrice;
            p.m_fImplied_Spread = instrument.m_fImpliedSpread;
            p.m_fFxRate = instrument.m_fFxRate;

            return p;
        }
        public Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, CBondDebugDataList>>>> GetDebugDataMap_Bonds(CBondDebugDataList PositionList)
        {
            // Scope || Portfolio ID | Security ID | Security ID LL | CBondDebugDataList
            Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, CBondDebugDataList>>>> pMap
                = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, CBondDebugDataList>>>>();
            Dictionary < string,Dictionary<string, Dictionary<string, CBondDebugDataList>>> scopeMap;
            Dictionary<string, Dictionary<string, CBondDebugDataList>> portfolioMap;
            Dictionary<string, CBondDebugDataList> securityMap;
            CBondDebugDataList securitiesList;
            foreach (CBondDebugData p in PositionList)
            {
                string scope = p.m_sScope3;
                if (pMap.ContainsKey(scope))
                {
                    scopeMap = pMap[scope];
                }
                else
                {
                    scopeMap = new Dictionary<string, Dictionary<string, Dictionary<string, CBondDebugDataList>>>();
                    pMap.Add(scope, scopeMap);
                }
                string portfolio = p.m_sPortfolio_ID;
                if (scopeMap.ContainsKey(portfolio))
                {
                    portfolioMap = scopeMap[portfolio];
                }
                else
                {
                    portfolioMap = new Dictionary<string, Dictionary<string, CBondDebugDataList>>();
                    scopeMap.Add(portfolio, portfolioMap);
                }
                string securityID = p.m_sSecurity_ID;
                if (portfolioMap.ContainsKey(securityID))
                {
                    securityMap = portfolioMap[securityID];
                }
                else
                {
                    securityMap = new Dictionary<string, CBondDebugDataList>();
                    portfolioMap.Add(securityID, securityMap);
                }
                string securityID_LL = p.m_sSecurity_ID_LL;
                if (securityMap.ContainsKey(securityID_LL))
                {
                    securitiesList = securityMap[securityID_LL];
                }
                else
                {
                    securitiesList = new CBondDebugDataList();
                    securityMap[securityID_LL] = securitiesList;
                }
                securitiesList.Add(p);
            }
            return pMap;
        }
        public CBondDebugData getLinkedPosition_Bond(CBondDebugData p_B)
        {
            CBondDebugData p_A = null;
            if (null != m_Scope_PortID_SecID_SecIDLL_BondDebugList_A)
            {
                if (m_Scope_PortID_SecID_SecIDLL_BondDebugList_A.ContainsKey(p_B.m_sScope3))
                {
                    Dictionary<string, Dictionary<string, Dictionary<string, CBondDebugDataList>>> s = m_Scope_PortID_SecID_SecIDLL_BondDebugList_A[p_B.m_sScope3];
                    if (s.ContainsKey(p_B.m_sPortfolio_ID))
                    {
                        Dictionary<string, Dictionary<string, CBondDebugDataList>> pf = s[p_B.m_sPortfolio_ID];
                        if (pf.ContainsKey(p_B.m_sSecurity_ID))
                        {
                            Dictionary<string, CBondDebugDataList> pSec = pf[p_B.m_sSecurity_ID];
                            if (pSec.ContainsKey(p_B.m_sSecurity_ID_LL))
                            {
                                CBondDebugDataList p = pSec[p_B.m_sSecurity_ID_LL];
                                p_A = p[0];
                                p_A.m_dNumberOfLinks++;
                                p_A.m_sRow_Debug_Linked = p_A.m_sRow_Debug;
                            }
                        }

                    }

                }
            }
            return p_A;
        }
        public void Process_Bond(FormMain frm, string name, DateTime dtReport, PositionList positions,
            ScenarioList scenarios, ErrorList errors, bool isCashFlowNeeded)
        {
            string sPrefix_PervePeriod = "";
            int cashflowCount = 0;
            object[,] expectedCashflowValues_RiskRente = null;
            object[,] expectedCashflowValues_RiskNeutral = null;
            object[,] expectedCashflowValues_RiskRente_EUR = null;
            object[,] expectedCashflowValues_RiskNeutral_EUR = null;
            List<object[]> OrigCashflowValues_RiskRente_EUR = null;
            List<object[]> OrigCashflowValues_RiskNeutral_EUR = null;
            List<object[]> OrigCashflowValues_RiskRente = null;
            List<object[]> OrigCashflowValues_RiskNeutral = null;
            try
            {
                List<string> headerNames = new List<string>();
                headerNames.Add("Period");
                headerNames.Add("Report Date");
                headerNames.Add("DATA SOURCE");
                headerNames.Add("Row ID Source");
                headerNames.Add("Row ID Debug");
                headerNames.Add("Period Prev");
                headerNames.Add("Report Date Prev");
                headerNames.Add("Row ID Debug Linked");
                headerNames.Add("Number of Links");
                headerNames.Add("Nominal_Value Prev");
                headerNames.Add("Implied_Spread Prev");
                headerNames.Add("FX Rate Prev");
                headerNames.Add("PositionId");
                headerNames.Add("Scope3");
                headerNames.Add("Is Look-Through Data");
                headerNames.Add("CurrencyHedgePerc");
                headerNames.Add("PortfolioId");
                headerNames.Add("SecurityID");
                headerNames.Add("SecurityName");
                headerNames.Add("SecurityID_LL");
                headerNames.Add("SecurityName_LL");
                headerNames.Add("InstrumentType");
                headerNames.Add("Bond_Type");
                headerNames.Add("DNB_Type");
                headerNames.Add("DNB_CountryUnionCode");
                headerNames.Add("DNB_IsFinancial");
                headerNames.Add("DNB_Rating");
                headerNames.Add("DNB_SpreadShock");
                headerNames.Add("CIC");
                headerNames.Add("CIC_LL");
                headerNames.Add("CouponType");
                headerNames.Add("Coupon");
                headerNames.Add("Coupon_Reference_Rate");
                headerNames.Add("Coupon_Spread");
                headerNames.Add("Coupon_Frequency");
                headerNames.Add("MaturityDate");
                headerNames.Add("Currency");
                headerNames.Add("FX Rate");
                headerNames.Add("In_Default");
                headerNames.Add("Nominal_Value");
                headerNames.Add("Market_Value");
                headerNames.Add("Model_Value");
                headerNames.Add("Implied_Spread");
                headerNames.Add("Implied_Spread_Found");
                headerNames.Add("Iterations");
                headerNames.Add("Duration");
                headerNames.Add("Maturity_in_Years");
                headerNames.Add("Value_with_Zero_Spread");
                headerNames.Add("CQS");
                headerNames.Add("Issuer_Group_Name");
                headerNames.Add("Issuer_Group_LEI");
                headerNames.Add("CollateralCoveragePerc");
                headerNames.Add("SCRLevel1Type");
                headerNames.Add("ASRLevel2Type");
                headerNames.Add("SCRStress");
                headerNames.Add("Spread_Duration");
                headerNames.Add("Message");
                double fDNB_SpreadShock = 0;
                object[,] scenarioValues = new object[positions.Count + 1, scenarios.Count];
                object[,] debugValues = new object[positions.Count + 1, headerNames.Count];
                int[] maximumNumberOfCF = new int[] { 0, 0 };
                if (isCashFlowNeeded)
                {
                    cashflowCount = m_dStandardizedCashFlow_DatePoints.Length;
                    expectedCashflowValues_RiskRente = new object[positions.Count + 1, cashflowCount + 1];
                    expectedCashflowValues_RiskNeutral = new object[positions.Count + 1, cashflowCount + 1];
                    expectedCashflowValues_RiskRente[0, 0] = "MarketValue";
                    expectedCashflowValues_RiskNeutral[0, 0] = "MarketValue";
                    expectedCashflowValues_RiskRente_EUR = new object[positions.Count + 1, cashflowCount + 1];
                    expectedCashflowValues_RiskNeutral_EUR = new object[positions.Count + 1, cashflowCount + 1];
                    expectedCashflowValues_RiskRente_EUR[0, 0] = "MarketValue";
                    expectedCashflowValues_RiskNeutral_EUR[0, 0] = "MarketValue";
                    for (int idxCashflow = 0; idxCashflow < cashflowCount; idxCashflow++)
                    {
                        DateTime dateHeader = m_dStandardizedCashFlow_DatePoints[idxCashflow];
                        expectedCashflowValues_RiskRente[0, idxCashflow + 1] = dateHeader;
                        expectedCashflowValues_RiskNeutral[0, idxCashflow + 1] = dateHeader;
                        expectedCashflowValues_RiskRente_EUR[0, idxCashflow + 1] = dateHeader;
                        expectedCashflowValues_RiskNeutral_EUR[0, idxCashflow + 1] = dateHeader;
                    }
                    object[] headerRow = new object[2];
                    headerRow[0] = "MarketValue";
                    headerRow[1] = "number of CF";
                    OrigCashflowValues_RiskRente_EUR = new List<object[]>();
                    OrigCashflowValues_RiskNeutral_EUR = new List<object[]>();
                    OrigCashflowValues_RiskRente_EUR.Add(headerRow);
                    OrigCashflowValues_RiskNeutral_EUR.Add(headerRow);
                    OrigCashflowValues_RiskRente = new List<object[]>();
                    OrigCashflowValues_RiskNeutral = new List<object[]>();
                    OrigCashflowValues_RiskRente.Add(headerRow);
                    OrigCashflowValues_RiskNeutral.Add(headerRow);
                }

                DateTime startTime = DateTime.Now;
                if (errors.CountErrors() == 0)
                {
                    int colDebug = 0;
                    for (colDebug = 0; colDebug < headerNames.Count; colDebug++)
                    {
                        debugValues[0, colDebug] = "'" + headerNames[colDebug];
                    }

                    // Initialize progress bar
                    frm.InitProgBar(scenarios.Count);

                    Curve scenarioZeroCurve;
                    string ccy;
                    Scenario scenario;
                    IndexCPI CPIindex;
                    double extraCreditSpread_SCRCharged;
                    double extraCreditSpread_Governmanets;
                    // Calculate market value for all combinations of positions/scenarios
                    for (int idxScenario = 0; idxScenario < scenarios.Count; idxScenario++)
                    {
                        scenario = scenarios[idxScenario];
                        scenarioValues[0, idxScenario] = scenario.m_sName;
                        frm.SetStatus("Verwerking " + name + " scenario " + scenario.m_sName);
                        CurveList scenarioZeroCurves = new CurveList();
                        foreach (ScenarioCurve scenarioCurve in scenario.m_YieldCurves)
                        {
                            ccy = scenarioCurve.m_sName;
                            scenarioZeroCurve = scenarioCurve.m_Curve;
                            scenarioZeroCurves.Add(ccy.ToUpper(), scenarioZeroCurve);
                        }
                        CurveList scenarioZeroInflationCurves = new CurveList();
                        SortedList<string, IndexCPI> scenarioIndexCPI_List = new SortedList<string, IndexCPI>();
                        foreach (ScenarioCurve scenarioCurve in scenario.m_InflationCurves)
                        {
                            ccy = scenarioCurve.m_sName;
                            scenarioZeroInflationCurves.Add(ccy.ToUpper(), scenarioCurve.m_Curve);
                            CPIindex = new IndexCPI();
                            CPIindex.SetInflationInstance(dtReport, 100, scenarioCurve.m_Curve, 100);
                            scenarioIndexCPI_List.Add(ccy.ToUpper(), CPIindex);
                        }

                        extraCreditSpread_SCRCharged = scenario.GetExtraCreditSpread("SCRCharged");
                        extraCreditSpread_Governmanets = scenario.GetExtraCreditSpread("Governments");
                        if (extraCreditSpread_Governmanets > 0)
                        {
                            extraCreditSpread_Governmanets += 0;
                        }
                        // Process all positions
                        for (int idxPosition = 0; idxPosition < positions.Count; idxPosition++)
                        {
                            int row = idxPosition + 1;
                            Position position = positions[idxPosition];
                            Instrument_Bond instrument = (Instrument_Bond)position.m_Instrument;
                            double fMaturity = instrument.getMaturity(dtReport);
                            // Create scenario discount curve
                            ccy = Position.TranslateCurrency_Curve(scenarioZeroCurves, position.m_sCurrency);
                            scenarioZeroCurve = scenarioZeroCurves[ccy];
                            // Create scenarion CPI index
                            CPIindex = null;
                            if (scenarioIndexCPI_List.Count > 0)
                            {
                                if (scenarioIndexCPI_List.ContainsKey(ccy))
                                {
                                    CPIindex = scenarioIndexCPI_List[ccy];
                                }
                                else
                                {
                                    CPIindex = null;
                                }
                            }
                            // Create scenario FX rate
                            ccy = Position.TranslateCurrency_Fx(scenario, position.m_sCurrency);
                            ScenarioValue scenarioValue = scenario.m_Fx.ByName(ccy);
                            double fxLevel = scenarioValue.m_fShockValue;
                            // Currency hedged funds:
                            double fxHedgePerc = 0;
                            if (m_CurrencyHedgePercentage_Fund.ContainsKey(position.m_sSecurityID))
                            {
                                fxHedgePerc = m_CurrencyHedgePercentage_Fund[position.m_sSecurityID];
                            }

                            // Calculate market value for this position with given scenario curve
                            double modelValueScenario = 0;
                            double modelValueScenario_fxBase;
                            double fImpliedSpread = instrument.m_fImpliedSpread;
                            double fFxRate = instrument.m_fFxRate;
                            if (null != instrument.m_BondDebugData_PrevPeriod)
                            {
                                if (scenario.usePrevPeriodStaticData())
                                {
                                    fImpliedSpread = instrument.m_BondDebugData_PrevPeriod.m_fImplied_Spread;
                                    fFxRate = instrument.m_BondDebugData_PrevPeriod.m_fFxRate;
                                }
                            }
                            fDNB_SpreadShock = 0;
                            // Check for DNB shock:
                            if (null != m_StressShocks_DNB)
                            {
                                if (instrument.m_OriginalBond.m_DNB_Type == DNB_Bond.Government)
                                {
                                    string sCountryCode = instrument.m_OriginalBond.m_sDNB_CountryUnion;
                                    if (m_StressShocks_DNB.m_GovermentBondsShocks.ContainsKey(sCountryCode))
                                    {
                                        fDNB_SpreadShock += getDNBShock(fMaturity, m_StressShocks_DNB.m_fGovermentBondsMatirities, m_StressShocks_DNB.m_GovermentBondsShocks[sCountryCode]);
                                    }
                                    fImpliedSpread += fDNB_SpreadShock;
                                }
                                else if (instrument.m_OriginalBond.m_DNB_Type == DNB_Bond.Corporate)
                                {
                                    string sCountryCode = instrument.m_OriginalBond.m_sDNB_CountryUnion;
                                    bool isFinancial = instrument.m_OriginalBond.m_bDNB_Financial;
                                    string sRating = instrument.m_OriginalBond.m_sDNB_Rating;
                                    Dictionary<string, Dictionary<string, double>> table = isFinancial ? m_StressShocks_DNB.m_CorporateBondsShocks_Financial : m_StressShocks_DNB.m_CorporateBondsShocks_nonFinancial;
                                    if (table.ContainsKey(sCountryCode))
                                    {
                                        if (table[sCountryCode].ContainsKey(sRating))
                                        {
                                            fDNB_SpreadShock = table[sCountryCode][sRating];
                                        }
                                    }
                                    fImpliedSpread += fDNB_SpreadShock;
                                }
                                else if (instrument.m_OriginalBond.m_DNB_Type == DNB_Bond.Covered)
                                {
                                    string sCountryCode = instrument.m_OriginalBond.m_sDNB_CountryUnion;
                                    string sRating = instrument.m_OriginalBond.m_sDNB_Rating;
                                    Dictionary<string, Dictionary<string, double>> table = m_StressShocks_DNB.m_CoveredBondsShocks;
                                    if (table.ContainsKey(sCountryCode))
                                    {
                                        if (table[sCountryCode].ContainsKey(sRating))
                                        {
                                            fDNB_SpreadShock = table[sCountryCode][sRating];
                                        }
                                    }
                                    fImpliedSpread += fDNB_SpreadShock;
                                }
                            }
                            if (null != position.m_SpreadRiskData)
                            {
                                if (position.m_SpreadRiskData.m_bSCRStress)
                                {
                                    modelValueScenario = instrument.getPrice(dtReport, scenarioZeroCurve, CPIindex, fImpliedSpread, fFxRate * fxLevel);
                                    modelValueScenario *= (1 - position.m_SpreadRiskData.m_fSpreadDuration * extraCreditSpread_SCRCharged);
                                }
                                else if ("Bond" == position.m_SpreadRiskData.m_sSCRLevel1Type &&
                                    "Bond_Government" == position.m_SpreadRiskData.m_sASRLevel2Type)
                                {
                                    modelValueScenario = instrument.getPrice(dtReport, scenarioZeroCurve, CPIindex, fImpliedSpread + extraCreditSpread_Governmanets, fFxRate * fxLevel);
                                }
                                else
                                {
                                    modelValueScenario = instrument.getPrice(dtReport, scenarioZeroCurve, CPIindex, fImpliedSpread, fFxRate * fxLevel);
                                }
                            }
                            else
                            {
                                modelValueScenario = instrument.getPrice(dtReport, scenarioZeroCurve, CPIindex, fImpliedSpread, fFxRate * fxLevel);
                            }
                            if (1 != fxLevel)
                            {
                                if (null != position.m_SpreadRiskData)
                                {
                                    if (position.m_SpreadRiskData.m_bSCRStress)
                                    {
                                        modelValueScenario_fxBase = instrument.getPrice(dtReport, scenarioZeroCurve, CPIindex, fImpliedSpread, fFxRate);
                                        modelValueScenario_fxBase *= (1 - position.m_SpreadRiskData.m_fSpreadDuration * extraCreditSpread_SCRCharged);
                                    }
                                    else if ("Bond" == position.m_SpreadRiskData.m_sSCRLevel1Type &&
                                        "Bond_Government" == position.m_SpreadRiskData.m_sASRLevel2Type)
                                    {
                                        modelValueScenario_fxBase = instrument.getPrice(dtReport, scenarioZeroCurve, CPIindex, fImpliedSpread + extraCreditSpread_Governmanets, fFxRate);
                                    }
                                    else
                                    {
                                        modelValueScenario_fxBase = instrument.getPrice(dtReport, scenarioZeroCurve, CPIindex, fImpliedSpread, fFxRate);
                                    }
                                }
                                else
                                {
                                    modelValueScenario_fxBase = instrument.getPrice(dtReport, scenarioZeroCurve, CPIindex, fImpliedSpread, fFxRate);
                                }
                            }
                            else
                            {
                                modelValueScenario_fxBase = modelValueScenario;
                            }
                            modelValueScenario = modelValueScenario_fxBase + (1 - fxHedgePerc)*(modelValueScenario - modelValueScenario_fxBase);
                            double scenarioDeltaValue = modelValueScenario - instrument.m_fModelPrice;
                            double marketValueScenario = instrument.m_fMarketPrice  + scenarioDeltaValue;
                            scenarioValues[row, idxScenario] = marketValueScenario;
                            if (scenario.isFairValueScenario())
                            {
                                instrument.m_BondDebugData_CurrentPeriod.m_sRow_Debug = (row+1).ToString();
                                colDebug = 0;
                                debugValues[row, colDebug++] = "'" + frm.m_sReportingPeriod;
                                debugValues[row, colDebug++] = instrument.m_OriginalBond.m_dtReport;
                                debugValues[row, colDebug++] = "'" + position.m_sDATA_Source;
                                debugValues[row, colDebug++] = "'" + position.m_sRow;
                                debugValues[row, colDebug++] = "'" + instrument.m_BondDebugData_CurrentPeriod.m_sRow_Debug;
                                if (null != instrument.m_BondDebugData_PrevPeriod)
                                {
                                    if ("" == sPrefix_PervePeriod) sPrefix_PervePeriod = instrument.m_BondDebugData_PrevPeriod.m_sReportingPeriod;
                                    debugValues[row, colDebug++] = "'" + instrument.m_BondDebugData_PrevPeriod.m_sReportingPeriod;
                                    debugValues[row, colDebug++] = instrument.m_BondDebugData_PrevPeriod.m_ReportDate;
                                    debugValues[row, colDebug++] = "'" + instrument.m_BondDebugData_PrevPeriod.m_sRow_Debug;
                                    debugValues[row, colDebug++] = "'" + instrument.m_BondDebugData_PrevPeriod.m_dNumberOfLinks;
                                    debugValues[row, colDebug++] = instrument.m_BondDebugData_PrevPeriod.m_fNominal_Value;
                                    debugValues[row, colDebug++] = instrument.m_BondDebugData_PrevPeriod.m_fImplied_Spread;
                                    debugValues[row, colDebug++] = instrument.m_BondDebugData_PrevPeriod.m_fFxRate;
                                }
                                else
                                {
                                    debugValues[row, colDebug++] = "";
                                    debugValues[row, colDebug++] = "";
                                    debugValues[row, colDebug++] = "";
                                    debugValues[row, colDebug++] = "";
                                    debugValues[row, colDebug++] = "";
                                    debugValues[row, colDebug++] = "";
                                    debugValues[row, colDebug++] = "";
                                }
                                debugValues[row, colDebug++] = "'" + position.m_sUniquePositionId;
                                debugValues[row, colDebug++] = "'" + position.m_sScope3;
                                debugValues[row, colDebug++] = instrument.m_bLookthroughBond;
                                debugValues[row, colDebug++] = fxHedgePerc;
                                debugValues[row, colDebug++] = "'" + position.m_sPortfolioId;
                                debugValues[row, colDebug++] = "'" + position.m_sSecurityID;
                                debugValues[row, colDebug++] = "'" + position.m_sSecurityName;
                                debugValues[row, colDebug++] = "'" + position.m_sSecurityID_LL;
                                debugValues[row, colDebug++] = "'" + position.m_sSecurityName_LL;
                                debugValues[row, colDebug++] = "'" + position.m_sSecurityType_LL;
                                debugValues[row, colDebug++] = "'" + instrument.m_OriginalBond.m_BondType.ToString();
                                debugValues[row, colDebug++] = "'" + instrument.m_OriginalBond.m_DNB_Type.ToString();
                                debugValues[row, colDebug++] = "'" + instrument.m_OriginalBond.m_sDNB_CountryUnion;
                                debugValues[row, colDebug++] = instrument.m_OriginalBond.m_bDNB_Financial;
                                debugValues[row, colDebug++] = "'" + instrument.m_OriginalBond.m_sDNB_Rating;
                                debugValues[row, colDebug++] = fDNB_SpreadShock;
                                debugValues[row, colDebug++] = "'" + position.m_sCIC;
                                debugValues[row, colDebug++] = "'" + position.m_sCIC_LL;
                                debugValues[row, colDebug++] = "'" + instrument.m_OriginalBond.m_sCouponType;
                                debugValues[row, colDebug++] = instrument.m_OriginalBond.m_fCoupon;
                                debugValues[row, colDebug++] = instrument.m_OriginalBond.m_fCouponReferenceRate;
                                debugValues[row, colDebug++] = instrument.m_OriginalBond.m_fCouponSpread;
                                debugValues[row, colDebug++] = instrument.m_OriginalBond.m_dCouponFrequency;
                                debugValues[row, colDebug++] = instrument.m_MaturityDate;
                                debugValues[row, colDebug++] = instrument.m_OriginalBond.m_sCurrency;
                                debugValues[row, colDebug++] = instrument.m_fFxRate;
                                debugValues[row, colDebug++] = instrument.m_bDefaulted;
                                debugValues[row, colDebug++] = instrument.m_fNominal;
                                debugValues[row, colDebug++] = instrument.m_fMarketPrice;
                                debugValues[row, colDebug++] = instrument.m_fModelPrice;
                                debugValues[row, colDebug++] = instrument.m_fImpliedSpread;
                                debugValues[row, colDebug++] = instrument.m_bImpliedSpreadFound;
                                debugValues[row, colDebug++] = instrument.m_dIterations;
                                debugValues[row, colDebug++] = instrument.m_fDuration;
                                debugValues[row, colDebug++] = fMaturity;
                                debugValues[row, colDebug++] = instrument.m_fModelValueAtZeroSpread;
                                debugValues[row, colDebug++] = "'" + position.m_sSecurityCreditQuality;
                                debugValues[row, colDebug++] = "'" + instrument.m_OriginalBond.m_sGroupCounterpartyName;
                                debugValues[row, colDebug++] = "'" + instrument.m_OriginalBond.m_sGroupCounterpartyLEI;
                                debugValues[row, colDebug++] = position.m_fCollateralCoveragePercentage;
                                if (null != position.m_SpreadRiskData)
                                {
                                    debugValues[row, colDebug++] = "'" + position.m_SpreadRiskData.m_sSCRLevel1Type;
                                    debugValues[row, colDebug++] = "'" + position.m_SpreadRiskData.m_sASRLevel2Type;
                                    debugValues[row, colDebug++] = "'" + position.m_SpreadRiskData.m_bSCRStress;
                                    debugValues[row, colDebug++] = position.m_SpreadRiskData.m_fSpreadDuration;
                                }
                                else
                                {
                                    debugValues[row, colDebug++] = "";
                                    debugValues[row, colDebug++] = "";
                                    debugValues[row, colDebug++] = "";
                                    debugValues[row, colDebug++] = "";
                                }
                                debugValues[row, colDebug++] = "'" + position.m_sMessage;
                                position.m_fFairValue = marketValueScenario;

                                if (isCashFlowNeeded)
                                {
                                    Curve scenarioZeroCurveEUR = scenarioZeroCurves["EUR"];
                                    instrument.calculateStandardizedCashFlows(dtReport, scenarioZeroCurve, scenarioZeroCurveEUR, m_dStandardizedCashFlow_DatePoints, CashFlowType.RiskRente);
                                    instrument.calculateStandardizedCashFlows(dtReport, scenarioZeroCurve, scenarioZeroCurveEUR, m_dStandardizedCashFlow_DatePoints, CashFlowType.RiskNeutral);
                                    instrument.calculateCashFlows_Orig_ToPrint(dtReport, scenarioZeroCurve, scenarioZeroCurveEUR, CashFlowType.RiskRente);
                                    instrument.calculateCashFlows_Orig_ToPrint(dtReport, scenarioZeroCurve, scenarioZeroCurveEUR, CashFlowType.RiskNeutral);
                                    // Standardtased Cash Flow:
                                    double[] CF_RiskRente = instrument.m_CashFlow_Reported[(int)CashFlowType.RiskRente];
                                    double[] CF_RiskNeutral = instrument.m_CashFlow_Reported[(int)CashFlowType.RiskNeutral];

                                    expectedCashflowValues_RiskRente[row, 0] = marketValueScenario;
                                    expectedCashflowValues_RiskNeutral[row, 0] = marketValueScenario;
                                    expectedCashflowValues_RiskRente_EUR[row, 0] = marketValueScenario;
                                    expectedCashflowValues_RiskNeutral_EUR[row, 0] = marketValueScenario;
                                    int nCF = cashflowCount;
                                    for (int idxCashflow = 0; idxCashflow < nCF; idxCashflow++)
                                    {
                                        expectedCashflowValues_RiskRente[row, idxCashflow + 1] = CF_RiskRente[idxCashflow]
                                            * instrument.m_fNominal;
                                        expectedCashflowValues_RiskNeutral[row, idxCashflow + 1] = CF_RiskNeutral[idxCashflow]
                                            * instrument.m_fNominal;
                                        expectedCashflowValues_RiskRente_EUR[row, idxCashflow + 1] = CF_RiskRente[idxCashflow]
                                            * instrument.m_fNominal * instrument.m_CashFlow_FXRate_Reported[idxCashflow];
                                        expectedCashflowValues_RiskNeutral_EUR[row, idxCashflow + 1] = CF_RiskNeutral[idxCashflow]
                                            * instrument.m_fNominal * instrument.m_CashFlow_FXRate_Reported[idxCashflow];
                                    }
                                    // Original Cash Flow:
                                    DateTime[] CF_Date_RiskRente = instrument.m_CashFlow_Orig_Date_ToPrint[(int)CashFlowType.RiskRente];
                                    DateTime[] CF_Date_RiskNeutral = instrument.m_CashFlow_Orig_Date_ToPrint[(int)CashFlowType.RiskNeutral];
                                    double[] CF_Amonut_RiskRente = instrument.m_CashFlow_Orig_Amonut_ToPrint[(int)CashFlowType.RiskRente];
                                    double[] CF_Amount_RiskNeutral = instrument.m_CashFlow_Orig_Amonut_ToPrint[(int)CashFlowType.RiskNeutral];
                                    double[] CF_FX_RiskRente = instrument.m_CashFlow_FXRate_Orig_ToPrint[(int)CashFlowType.RiskRente];
                                    double[] CF_FX_RiskNeutral = instrument.m_CashFlow_FXRate_Orig_ToPrint[(int)CashFlowType.RiskNeutral];

                                    nCF = CF_Date_RiskRente.Length;
                                    if (nCF > maximumNumberOfCF[(int)CashFlowType.RiskRente])
                                    {
                                        maximumNumberOfCF[(int)CashFlowType.RiskRente] = nCF;
                                    }
                                    object[] origCashflowValues_RiskRente_EUR = new object[2 + 2 * nCF];
                                    object[] origCashflowValues_RiskRente = new object[2 + 2 * nCF];
                                    origCashflowValues_RiskRente_EUR[0] = marketValueScenario;
                                    origCashflowValues_RiskRente_EUR[1] = nCF;
                                    origCashflowValues_RiskRente[0] = marketValueScenario;
                                    origCashflowValues_RiskRente[1] = nCF;
                                    for (int idxCashflow = 0; idxCashflow < nCF; idxCashflow++)
                                    {
                                        origCashflowValues_RiskRente_EUR[idxCashflow + 2] = CF_Date_RiskRente[idxCashflow];
                                        origCashflowValues_RiskRente_EUR[idxCashflow + 2 + nCF] = instrument.m_fNominal *
                                            CF_Amonut_RiskRente[idxCashflow] * CF_FX_RiskRente[idxCashflow];
                                        origCashflowValues_RiskRente[idxCashflow + 2] = CF_Date_RiskRente[idxCashflow];
                                        origCashflowValues_RiskRente[idxCashflow + 2 + nCF] = instrument.m_fNominal *
                                            CF_Amonut_RiskRente[idxCashflow];
                                    }
                                    OrigCashflowValues_RiskRente_EUR.Add(origCashflowValues_RiskRente_EUR);
                                    OrigCashflowValues_RiskRente.Add(origCashflowValues_RiskRente);

                                    nCF = CF_Date_RiskNeutral.Length;
                                    if (nCF > maximumNumberOfCF[(int)CashFlowType.RiskNeutral])
                                    {
                                        maximumNumberOfCF[(int)CashFlowType.RiskNeutral] = nCF;
                                    }
                                    object[] origCashflowValues_RiskNeutral_EUR = new object[2 + 2 * nCF];
                                    origCashflowValues_RiskNeutral_EUR[0] = marketValueScenario;
                                    origCashflowValues_RiskNeutral_EUR[1] = nCF;
                                    object[] origCashflowValues_RiskNeutral = new object[2 + 2 * nCF];
                                    origCashflowValues_RiskNeutral[0] = marketValueScenario;
                                    origCashflowValues_RiskNeutral[1] = nCF;
                                    for (int idxCashflow = 0; idxCashflow < nCF; idxCashflow++)
                                    {
                                        origCashflowValues_RiskNeutral_EUR[idxCashflow + 2] = CF_Date_RiskNeutral[idxCashflow];
                                        origCashflowValues_RiskNeutral_EUR[idxCashflow + 2 + nCF] = instrument.m_fNominal *
                                            CF_Amount_RiskNeutral[idxCashflow] * CF_FX_RiskNeutral[idxCashflow];
                                        origCashflowValues_RiskNeutral[idxCashflow + 2] = CF_Date_RiskNeutral[idxCashflow];
                                        origCashflowValues_RiskNeutral[idxCashflow + 2 + nCF] = instrument.m_fNominal *
                                            CF_Amount_RiskNeutral[idxCashflow];
                                    }
                                    OrigCashflowValues_RiskNeutral_EUR.Add(origCashflowValues_RiskNeutral_EUR);
                                    OrigCashflowValues_RiskNeutral.Add(origCashflowValues_RiskNeutral);
                                }

                                System.Windows.Forms.Application.DoEvents();
                            }
                        }
                        // Update progress bar
                        frm.IncrementProgBar();
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                DateTime endTime = DateTime.Now;
                double runTime = (endTime - startTime).TotalSeconds;

                // Write aggregated market values to results workbook
                frm.SetStatus("Opslaan resultaten " + name);
                Scenario baseScenario = scenarios.getScenarioFairValue();
                CurveList baseCurves = new CurveList();
                foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
                {
                    baseCurves.Add(scenarioCurve.m_sName.ToUpper(), scenarioCurve.m_Curve);
                }
                object[,] curveValues = baseCurves.ToArray(true);
                object[,] errorValues = errors.ToArray();
                if ("" != sPrefix_PervePeriod) sPrefix_PervePeriod = " vs " + sPrefix_PervePeriod + " - ";
                frm.WriteToExcel(name + sPrefix_PervePeriod, positions, scenarioValues, curveValues, debugValues, errorValues, runTime, null, null, false);
                if (isCashFlowNeeded)
                {
                    frm.WriteToExcel(name + " Cash flows Risk Neutral (EUR)", positions, expectedCashflowValues_RiskNeutral_EUR, curveValues, debugValues, errorValues, runTime, null, null, false);
                    frm.WriteToExcel(name + " Cash flows Risk Rente (EUR)", positions, expectedCashflowValues_RiskRente_EUR, curveValues, debugValues, errorValues, runTime, null, null, false);
                    frm.WriteToExcel(name + " Cash flows Risk Neutral", positions, expectedCashflowValues_RiskNeutral, curveValues, debugValues, errorValues, runTime, null, null, false);
                    frm.WriteToExcel(name + " Cash flows Risk Rente", positions, expectedCashflowValues_RiskRente, curveValues, debugValues, errorValues, runTime, null, null, false);
                    object[,] oOrigCashflowValues_RiskRente_EUR = new object[positions.Count + 1, 2 + 2 * maximumNumberOfCF[(int)CashFlowType.RiskRente]];
                    object[,] oOrigCashflowValues_RiskRente = new object[positions.Count + 1, 2 + 2 * maximumNumberOfCF[(int)CashFlowType.RiskRente]];
                    object[] header = OrigCashflowValues_RiskRente_EUR[0];
                    oOrigCashflowValues_RiskRente_EUR[0, 0] = header[0];
                    oOrigCashflowValues_RiskRente_EUR[0, 1] = header[1];
                    oOrigCashflowValues_RiskRente[0, 0] = header[0];
                    oOrigCashflowValues_RiskRente[0, 1] = header[1];
                    object[,] oOrigCashflowValues_RiskNeutral_EUR = new object[positions.Count + 1, 2 + 2 * maximumNumberOfCF[(int)CashFlowType.RiskNeutral]];
                    object[,] oOrigCashflowValues_RiskNeutral = new object[positions.Count + 1, 2 + 2 * maximumNumberOfCF[(int)CashFlowType.RiskNeutral]];
                    header = OrigCashflowValues_RiskNeutral_EUR[0];
                    oOrigCashflowValues_RiskNeutral_EUR[0, 0] = header[0];
                    oOrigCashflowValues_RiskNeutral_EUR[0, 1] = header[1];
                    oOrigCashflowValues_RiskNeutral[0, 0] = header[0];
                    oOrigCashflowValues_RiskNeutral[0, 1] = header[1];

                    object[][] aOrigCashflowValues_RiskRente_EUR = new object[positions.Count + 1][];
                    aOrigCashflowValues_RiskRente_EUR[0] = OrigCashflowValues_RiskRente_EUR[0];
                    object[][] aOrigCashflowValues_RiskNeutral_EUR = new object[positions.Count + 1][];
                    aOrigCashflowValues_RiskNeutral_EUR[0] = OrigCashflowValues_RiskNeutral_EUR[0];

                    object[][] aOrigCashflowValues_RiskRente = new object[positions.Count + 1][];
                    aOrigCashflowValues_RiskRente[0] = OrigCashflowValues_RiskRente[0];
                    object[][] aOrigCashflowValues_RiskNeutral = new object[positions.Count + 1][];
                    aOrigCashflowValues_RiskNeutral[0] = OrigCashflowValues_RiskNeutral[0];

                    for (int idxPosition = 0; idxPosition < positions.Count; idxPosition++)
                    {
                        int row = idxPosition + 1;
                        aOrigCashflowValues_RiskRente_EUR[idxPosition + 1] = OrigCashflowValues_RiskRente_EUR[idxPosition + 1];
                        aOrigCashflowValues_RiskNeutral_EUR[idxPosition + 1] = OrigCashflowValues_RiskNeutral_EUR[idxPosition + 1];
                        aOrigCashflowValues_RiskRente[idxPosition + 1] = OrigCashflowValues_RiskRente[idxPosition + 1];
                        aOrigCashflowValues_RiskNeutral[idxPosition + 1] = OrigCashflowValues_RiskNeutral[idxPosition + 1];
                        object[] rrCF_EUR = OrigCashflowValues_RiskRente_EUR[idxPosition + 1];
                        object[] rrCF = OrigCashflowValues_RiskRente[idxPosition + 1];
                        for (int i = 0; i < rrCF_EUR.Length; i++)
                        {
                            oOrigCashflowValues_RiskRente_EUR[row, i] = rrCF_EUR[i];
                            oOrigCashflowValues_RiskRente[row, i] = rrCF[i];
                        }
                        object[] rnCF_EUR = OrigCashflowValues_RiskNeutral_EUR[idxPosition + 1];
                        object[] rnCF = OrigCashflowValues_RiskNeutral[idxPosition + 1];
                        for (int i = 0; i < rnCF_EUR.Length; i++)
                        {
                            oOrigCashflowValues_RiskNeutral_EUR[row, i] = rnCF_EUR[i];
                            oOrigCashflowValues_RiskNeutral[row, i] = rnCF[i];
                        }
                    }
                    frm.WriteToExcel(name + " Cash flows Original Risk Rente", positions, oOrigCashflowValues_RiskRente, curveValues, debugValues, errorValues, runTime, null, null, false);
                    frm.WriteToExcel(name + " Cash flows Original Risk Neutral", positions, oOrigCashflowValues_RiskNeutral, curveValues, debugValues, errorValues, runTime, null, null, false);
                    frm.WriteToExcel(name + " Cash flows Original Risk Rente (EUR)", positions, oOrigCashflowValues_RiskRente_EUR, curveValues, debugValues, errorValues, runTime, null, null, false);
                    frm.WriteToExcel(name + " Cash flows Original Risk Neutral (EUR)", positions, oOrigCashflowValues_RiskNeutral_EUR, curveValues, debugValues, errorValues, runTime, null, null, false);
                }

                if (errors.CountErrors() > 0)
                {
                    MessageBox.Show("Fouten in verwerking " + name + ", bekijk Error werkblad in output bestand", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (IOException exc)
            {
                MessageBox.Show("Fout in verwerking " + name + ":\n" + exc.Message, "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            frm.HideProgBar();
        }
        /** CASH */
        public void Process_Cash(FormMain frm, string name, DateTime dtReport, PositionList positions,
            ScenarioList scenarios, ErrorList errors, bool isCashFlowNeeded)
        {
            try
            {
                List<string> headerNames = new List<string>();
                headerNames.Add("Period");
                headerNames.Add("BalanceType");
                headerNames.Add("Group");
                headerNames.Add("InstrumentType");
                headerNames.Add("DATA SOURCE");
                headerNames.Add("Row ID");
                headerNames.Add("PositionId");
                headerNames.Add("SelectieIndex_LL");
                headerNames.Add("Scope3");
                headerNames.Add("Scope Issuer");
                headerNames.Add("Scope Investor");
                headerNames.Add("ICO");
                headerNames.Add("Is Look-Through Data");
                headerNames.Add("CurrencyHedgePerc");
                headerNames.Add("PortfolioId");
                headerNames.Add("SecurityID");
                headerNames.Add("SecurityName");
                headerNames.Add("SecurityID_LL");
                headerNames.Add("SecurityName_LL");
                headerNames.Add("CIC");
                headerNames.Add("CIC_LL");
                headerNames.Add("Coupon");
                headerNames.Add("Start Date");
                headerNames.Add("End Date");
                headerNames.Add("Is Overnight");
                headerNames.Add("Currency");
                headerNames.Add("FX Rate");
                headerNames.Add("Nominal_Value");
                headerNames.Add("Market_Value");
                headerNames.Add("Moarket_Value (EUR)");
                headerNames.Add("Counterparty_Group_Name");
                headerNames.Add("Counterparty_Group_LEI");
                headerNames.Add("Counterparty_Group_CQS");
                headerNames.Add("CollateralCoveragePerc");
                headerNames.Add("Type");
                headerNames.Add("Message");


                object[,] scenarioValues = new object[positions.Count + 1, scenarios.Count];
                object[,] debugValues = new object[positions.Count + 1, headerNames.Count];

                DateTime startTime = DateTime.Now;
                if (errors.CountErrors() == 0)
                {
                    int colDebug = 0;
                    for (colDebug = 0; colDebug < headerNames.Count; colDebug++)
                    {
                        debugValues[0, colDebug] = "'" + headerNames[colDebug];
                    }

                    // Initialize progress bar
                    frm.InitProgBar(scenarios.Count);

                    Curve scenarioZeroCurve;
                    string ccy;
                    Scenario scenario;
                    // Calculate market value for all combinations of positions/scenarios
                    for (int idxScenario = 0; idxScenario < scenarios.Count; idxScenario++)
                    {
                        scenario = scenarios[idxScenario];
                        scenarioValues[0, idxScenario] = scenario.m_sName;
                        frm.SetStatus("Verwerking " + name + " scenario " + scenario.m_sName);
                        CurveList scenarioZeroCurves = new CurveList();
                        foreach (ScenarioCurve scenarioCurve in scenario.m_YieldCurves)
                        {
                            ccy = scenarioCurve.m_sName;
                            scenarioZeroCurve = scenarioCurve.m_Curve;
                            scenarioZeroCurves.Add(ccy.ToUpper(), scenarioZeroCurve);
                        }
                        // Process all positions
                        for (int idxPosition = 0; idxPosition < positions.Count; idxPosition++)
                        {
                            Position position = positions[idxPosition];
                            Instrument_Cash instrument = (Instrument_Cash)position.m_Instrument;

                            // Create scenario discount curve
                            ccy = Position.TranslateCurrency_Curve(scenarioZeroCurves, position.m_sCurrency);
                            scenarioZeroCurve = scenarioZeroCurves[ccy];
                            // Create scenario FX rate
                            ccy = Position.TranslateCurrency_Fx(scenario, position.m_sCurrency);
                            ScenarioValue scenarioValue = scenario.m_Fx.ByName(ccy);
                            double fxLevel = scenarioValue.m_fShockValue;
                            // Currency hedged funds:
                            double fxHedgePerc = 0;
                            if (m_CurrencyHedgePercentage_Fund.ContainsKey(position.m_sSecurityID))
                            {
                                fxHedgePerc = m_CurrencyHedgePercentage_Fund[position.m_sSecurityID];
                            }

                            // Calculate market value for this position with given scenario curve
                            double modelValueScenario_base = instrument.getPrice(dtReport, 1);
                            double modelValueScenario_shock = instrument.getPrice(dtReport, fxLevel);
                            double marketValueScenario = modelValueScenario_base + (1 - fxHedgePerc)*(modelValueScenario_shock - modelValueScenario_base);
                            scenarioValues[idxPosition + 1, idxScenario] = marketValueScenario;
                            if (scenario.isFairValueScenario())
                            {
                                colDebug = 0;
                                debugValues[idxPosition + 1, colDebug++] = "'" + frm.m_sReportingPeriod;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sBalanceType;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sGroup;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSecurityType_LL;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sDATA_Source;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sRow;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sUniquePositionId;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSelectieIndex_LL;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sScope3;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sScope3_Issuer;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sScope3_Investor;
                                debugValues[idxPosition + 1, colDebug++] = position.m_bICO;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_OriginalData.m_bLookThroughData;
                                debugValues[idxPosition + 1, colDebug++] = fxHedgePerc;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sPortfolioId;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSecurityID;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSecurityName;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSecurityID_LL;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSecurityName_LL;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCIC;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sCIC_LL;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_OriginalData.m_fCoupon;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_OriginalData.m_dtStartDate;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_OriginalData.m_dtEndDate;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_OriginalData.m_bIsOvernight;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_OriginalData.m_sCurrency;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fFxRate;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fNominal;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fMarketValue;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_OriginalData.m_fMarketValue_EUR;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_OriginalData.m_sGroupCounterpartyName;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_OriginalData.m_sGroupCounterpartyLEI;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_OriginalData.m_sGroupCounterpartyCQS;
                                debugValues[idxPosition + 1, colDebug++] = position.m_fCollateralCoveragePercentage;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_OriginalData.m_dType;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sMessage;
                                position.m_fFairValue = marketValueScenario;

                                System.Windows.Forms.Application.DoEvents();
                            }
                        }
                        // Update progress bar
                        frm.IncrementProgBar();
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                DateTime endTime = DateTime.Now;
                double runTime = (endTime - startTime).TotalSeconds;

                // Write aggregated market values to results workbook
                frm.SetStatus("Opslaan resultaten " + name);
                Scenario baseScenario = scenarios.getScenarioFairValue();
                CurveList baseCurves = new CurveList();
                foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
                {
                    baseCurves.Add(scenarioCurve.m_sName.ToUpper(), scenarioCurve.m_Curve);
                }
                object[,] curveValues = baseCurves.ToArray(true);
                object[,] errorValues = errors.ToArray();
                frm.WriteToExcel(name, positions, scenarioValues, curveValues, debugValues, errorValues, runTime, null, null, false);

                if (errors.CountErrors() > 0)
                {
                    MessageBox.Show("Fouten in verwerking " + name + ", bekijk Error werkblad in output bestand", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (IOException exc)
            {
                MessageBox.Show("Fout in verwerking " + name + ":\n" + exc.Message, "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            frm.HideProgBar();
        }
        public void Process_FI(FormMain frm, string name, DateTime dtReport, PositionList positions,
            ScenarioList scenarios, ErrorList errors, bool isCashFlowNeeded)
        {
            int cashflowCount = 0;
            object[,] expectedCashflowValues = null;
            try
            {
                object[,] scenarioValues = new object[positions.Count + 1, scenarios.Count];
                object[,] debugValues = new object[positions.Count + 1, 20];
                if (isCashFlowNeeded)
                {
                    cashflowCount = m_dStandardizedCashFlow_DatePoints.Length;
                    expectedCashflowValues = new object[positions.Count + 1, cashflowCount + 1];
                    expectedCashflowValues[0, 0] = "MarketValue";
                    for (int idxCashflow = 0; idxCashflow < cashflowCount; idxCashflow++)
                    {
                        DateTime dateHeader = m_dStandardizedCashFlow_DatePoints[idxCashflow];
                        expectedCashflowValues[0, idxCashflow + 1] = dateHeader;
                    }
                }

                DateTime startTime = DateTime.Now;
                if (errors.CountErrors() == 0)
                {
                    int colDebug = 0;
                    debugValues[0, colDebug++] = "PositionId";
                    debugValues[0, colDebug++] = "PortfolioId";
                    debugValues[0, colDebug++] = "SecurityName";
                    debugValues[0, colDebug++] = "InstrumentType";
                    debugValues[0, colDebug++] = "CouponType";
                    debugValues[0, colDebug++] = "MaturityDate";
                    debugValues[0, colDebug++] = "CleanValue";
                    debugValues[0, colDebug++] = "ImpliedSpread";
                    debugValues[0, colDebug++] = "Iterations";
                    debugValues[0, colDebug++] = "Duration";
                    debugValues[0, colDebug++] = "Value with Zero Spread";
                    debugValues[0, colDebug++] = "Rating";
                    debugValues[0, colDebug++] = "CollateralCoveragePerc";
                    debugValues[0, colDebug++] = "Scope3";
                    debugValues[0, colDebug++] = "Coupon";
                    debugValues[0, colDebug++] = "SCRLevel1Type";
                    debugValues[0, colDebug++] = "ASRLevel2Type";
                    debugValues[0, colDebug++] = "SCRStress";
                    debugValues[0, colDebug++] = "Spread Duration";
                    debugValues[0, colDebug++] = "Message";

                    // Initialize progress bar
                    frm.InitProgBar(scenarios.Count);

                    Curve scenarioZeroCurve;
                    string ccy;
                    Scenario scenario;
                    double extraCreditSpread_SCRCharged;
                    double extraCreditSpread_Governmanets;
                    // Calculate market value for all combinations of positions/scenarios
                    for (int idxScenario = 0; idxScenario < scenarios.Count; idxScenario++)
                    {
                        scenario = scenarios[idxScenario];
                        scenarioValues[0, idxScenario] = scenario.m_sName;
                        frm.SetStatus("Verwerking " + name + " scenario " + scenario.m_sName);
                        CurveList scenarioZeroCurves = new CurveList();
                        foreach (ScenarioCurve scenarioCurve in scenario.m_YieldCurves)
                        {
                            scenarioZeroCurve = scenarioCurve.m_Curve;
                            scenarioZeroCurves.Add(scenarioCurve.m_sName.ToUpper(), scenarioZeroCurve);
                        }
                        extraCreditSpread_SCRCharged = scenario.GetExtraCreditSpread("SCRCharged");
                        extraCreditSpread_Governmanets = scenario.GetExtraCreditSpread("Governments");
                        if (extraCreditSpread_Governmanets > 0)
                        {
                            extraCreditSpread_Governmanets += 0;
                        }
                        // Process all positions
                        for (int idxPosition = 0; idxPosition < positions.Count; idxPosition++)
                        {
                            Position position = positions[idxPosition];
                            Instrument_Cashflow instrument = (Instrument_Cashflow)position.m_Instrument;

                            // Create scenario discount curve
                            ccy = Position.TranslateCurrency_Curve(scenarioZeroCurves, position.m_sCurrency);
                            scenarioZeroCurve = scenarioZeroCurves[ccy];
                            // Create scenario FX rate
                            ccy = Position.TranslateCurrency_Fx(scenario, position.m_sCurrency);
                            ScenarioValue scenarioValue = scenario.m_Fx.ByName(ccy);
                            double fxLevel = scenarioValue.m_fShockValue; // it is not used because the Cash flow model works only with EUR cash flows

                            // Calculate market value for this position with given scenario curve
                            double modelValueScenario = 0;
                            if (null != position.m_SpreadRiskData)
                            {
                                if (position.m_SpreadRiskData.m_bSCRStress)
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread);
                                    modelValueScenario *= (1 - position.m_SpreadRiskData.m_fSpreadDuration * extraCreditSpread_SCRCharged);
                                }
                                else if ("Bond" == position.m_SpreadRiskData.m_sSCRLevel1Type &&
                                    "Bond_Government" == position.m_SpreadRiskData.m_sASRLevel2Type)
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread + extraCreditSpread_Governmanets);
                                }
                                else
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread);
                                }
                            }
                            else
                            {
                                if ("" != position.m_sCIC_LL && "A2" == position.m_sCIC_LL.Substring(2, 2))
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread + extraCreditSpread_Governmanets);
                                }
                                else
                                {
                                    modelValueScenario = instrument.MarketValue(dtReport, scenarioZeroCurve, instrument.m_fImpliedSpread);
                                }
                            }
                            double marketValueScenario = instrument.m_fDirtyValue + (modelValueScenario - instrument.m_fDirtyValue);
                            scenarioValues[idxPosition + 1, idxScenario] = modelValueScenario;
                            if (scenario.isFairValueScenario())
                            {
                                colDebug = 0;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sPositionId;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sPortfolioId;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sSecurityName_LL;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_sInstrumentType;
                                debugValues[idxPosition + 1, colDebug++] = "'" + instrument.m_sCouponType;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_MaturityDate;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fDirtyValue;
                                if (double.IsNaN(instrument.m_fImpliedSpread))
                                {
                                    debugValues[idxPosition + 1, colDebug++] = "#N/A";
                                }
                                else
                                {
                                    debugValues[idxPosition + 1, colDebug++] = instrument.m_fImpliedSpread;
                                }
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_dIterations;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fDuration;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fValueAtZeroSpread;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sRating;
                                debugValues[idxPosition + 1, colDebug++] = position.m_fCollateralCoveragePercentage;
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sScope3;
                                debugValues[idxPosition + 1, colDebug++] = instrument.m_fCoupon;
                                if (null != position.m_SpreadRiskData)
                                {
                                    debugValues[idxPosition + 1, colDebug++] = "'" + position.m_SpreadRiskData.m_sSCRLevel1Type;
                                    debugValues[idxPosition + 1, colDebug++] = "'" + position.m_SpreadRiskData.m_sASRLevel2Type;
                                    debugValues[idxPosition + 1, colDebug++] = "'" + position.m_SpreadRiskData.m_bSCRStress;
                                    debugValues[idxPosition + 1, colDebug++] = position.m_SpreadRiskData.m_fSpreadDuration;
                                }
                                else
                                {
                                    debugValues[idxPosition + 1, colDebug++] = "";
                                    debugValues[idxPosition + 1, colDebug++] = "";
                                    debugValues[idxPosition + 1, colDebug++] = "";
                                    debugValues[idxPosition + 1, colDebug++] = "";
                                }
                                debugValues[idxPosition + 1, colDebug++] = "'" + position.m_sMessage;
                                position.m_fFairValue = modelValueScenario;

                                if (isCashFlowNeeded)
                                {
                                    instrument.calculateStandardizedCashFlows(dtReport, m_dStandardizedCashFlow_DatePoints);
                                    expectedCashflowValues[idxPosition + 1, 0] = marketValueScenario;
                                    for (int idxCashflow = 0; idxCashflow < cashflowCount; idxCashflow++)
                                    {
                                        expectedCashflowValues[idxPosition + 1, idxCashflow + 1] = instrument.m_CashFlow_Reported[idxCashflow];
                                    }
                                }
                            }

                            System.Windows.Forms.Application.DoEvents();
                        }

                        // Update progress bar
                        frm.IncrementProgBar();
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                DateTime endTime = DateTime.Now;
                double runTime = (endTime - startTime).TotalSeconds;

                // Write aggregated market values to results workbook
                frm.SetStatus("Opslaan resultaten " + name);
                Scenario baseScenario = scenarios.getScenarioFairValue();
                CurveList baseCurves = new CurveList();
                foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
                {
                    baseCurves.Add(scenarioCurve.m_sName.ToUpper(), scenarioCurve.m_Curve);
                }
                object[,] curveValues = baseCurves.ToArray(true);
                object[,] errorValues = errors.ToArray();
                frm.WriteToExcel(name, positions, scenarioValues, curveValues, debugValues, errorValues, runTime, null, null, false);
                if (isCashFlowNeeded)
                {
                    frm.WriteToExcel(name + " Cash flows (EUR)", positions, expectedCashflowValues, curveValues, debugValues, errorValues, runTime, null, null, false);
                }

                if (errors.CountErrors() > 0)
                {
                    MessageBox.Show("Fouten in verwerking " + name + ", bekijk Error werkblad in output bestand", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (IOException exc)
            {
                MessageBox.Show("Fout in verwerking " + name + ":\n" + exc.Message, "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            frm.HideProgBar();
        }
        public void Process_Swap(FormMain frm, string name, DateTime? dtReport, PositionList positions, 
            ScenarioList scenarios, ErrorList errors, bool isCashFlowNeeded)
        {
            int cashflowCount = m_dStandardizedCashFlow_DatePoints.Length;
            List<string> debugValues_headers_List = new List<string>();
            debugValues_headers_List.Add("Period");
            debugValues_headers_List.Add("DATA SOURCE");
            debugValues_headers_List.Add("Row ID");
            debugValues_headers_List.Add("PositionId");
            debugValues_headers_List.Add("SecurityName");
            debugValues_headers_List.Add("Volume");
            debugValues_headers_List.Add("InstrumentType");
            debugValues_headers_List.Add("Payment Type");
            debugValues_headers_List.Add("Swap Type");
            debugValues_headers_List.Add("CURR Leg 1");
            debugValues_headers_List.Add("CURR Leg 2");
            debugValues_headers_List.Add("StartDate");
            debugValues_headers_List.Add("MaturityDate");
            debugValues_headers_List.Add("ID Leg 1");
            debugValues_headers_List.Add("ID Leg 2");
            debugValues_headers_List.Add("PaymentType Leg 1");
            debugValues_headers_List.Add("PaymentType Leg 2");
            debugValues_headers_List.Add("Is Fixed Leg 1");
            debugValues_headers_List.Add("Is Fixed Leg 2");
            debugValues_headers_List.Add("Payment Sign Leg 1");
            debugValues_headers_List.Add("Payment Sing Leg 2");
            debugValues_headers_List.Add("Volume Leg 1");
            debugValues_headers_List.Add("Volume Leg 2");
            debugValues_headers_List.Add("Rate Leg 1");
            debugValues_headers_List.Add("Rate Leg 2");
            debugValues_headers_List.Add("Frequency Leg 1");
            debugValues_headers_List.Add("Frequency Leg 2");
            debugValues_headers_List.Add("Expiry");
            debugValues_headers_List.Add("Tenor");
            debugValues_headers_List.Add("Model Type");
            debugValues_headers_List.Add("MarketPrice");
            debugValues_headers_List.Add("ModelPrice");
            debugValues_headers_List.Add("Duration");
            debugValues_headers_List.Add("MarketPrice Leg 1");
            debugValues_headers_List.Add("ModelPrice  Leg 1");
            debugValues_headers_List.Add("Duration Leg 1");
            debugValues_headers_List.Add("MarketPrice Leg 2");
            debugValues_headers_List.Add("ModelPrice  Leg 2");
            debugValues_headers_List.Add("Duration Leg 2");
            debugValues_headers_List.Add("Discounting Spread Leg 1");
            debugValues_headers_List.Add("Discounting Spread Leg 2");
            debugValues_headers_List.Add("Floating Rate Spread Leg 1");
            debugValues_headers_List.Add("Floating Rate Spread Leg 2");
            debugValues_headers_List.Add("FX Leg 1");
            debugValues_headers_List.Add("FX Leg 2");
            debugValues_headers_List.Add("Discount Rate Leg 1");
            debugValues_headers_List.Add("Discount Rate Leg 2");
            debugValues_headers_List.Add("Discount Spread Duration Leg 1");
            debugValues_headers_List.Add("Discount Spread Duration Leg 2");
            debugValues_headers_List.Add("Iterations");
            

            int positionIndex = -1;
            string sSecurutyID = "";
            try
            {
                object[,] scenarioValuesLegFixed = new object[positions.Count + 1, scenarios.Count];
                object[,] scenarioValuesLegFloat = new object[positions.Count + 1, scenarios.Count];
                object[,] scenarioValues = new object[positions.Count + 1, scenarios.Count];
                object[,] scenarioValuesBothLegs = new object[2*positions.Count + 1, scenarios.Count];
                object[,] debugValues = new object[positions.Count + 1, debugValues_headers_List.Count];
                object[,] debugValues2 = new object[2*positions.Count + 1, debugValues_headers_List.Count];
                object[,] expectedCashflowValues_RiskRente_EUR = new object[2 * positions.Count + 1, cashflowCount + 1];
                object[,] expectedCashflowValues_RiskNeutral_EUR = new object[2 * positions.Count + 1, cashflowCount + 1];
                object[,] expectedCouponsValues_RiskRente_EUR = new object[2 * positions.Count + 1, cashflowCount + 1];
                object[,] expectedCouponsValues_RiskNeutral_EUR = new object[2 * positions.Count + 1, cashflowCount + 1];
                object[,] expectedNominalValues_RiskRente_EUR = new object[2 * positions.Count + 1, cashflowCount + 1];
                object[,] expectedNominalValues_RiskNeutral_EUR = new object[2 * positions.Count + 1, cashflowCount + 1];
                expectedCashflowValues_RiskRente_EUR[0, 0] = "MarketValue";
                expectedCashflowValues_RiskNeutral_EUR[0, 0] = "MarketValue";
                expectedCouponsValues_RiskRente_EUR[0, 0] = "MarketValue";
                expectedCouponsValues_RiskNeutral_EUR[0, 0] = "MarketValue";
                expectedNominalValues_RiskRente_EUR[0, 0] = "MarketValue";
                expectedNominalValues_RiskNeutral_EUR[0, 0] = "MarketValue";
                object[,] expectedCashflowValues_RiskRente = new object[2 * positions.Count + 1, cashflowCount + 1];
                object[,] expectedCashflowValues_RiskNeutral = new object[2 * positions.Count + 1, cashflowCount + 1];
                object[,] expectedCouponsValues_RiskRente = new object[2 * positions.Count + 1, cashflowCount + 1];
                object[,] expectedCouponsValues_RiskNeutral = new object[2 * positions.Count + 1, cashflowCount + 1];
                object[,] expectedNominalValues_RiskRente = new object[2 * positions.Count + 1, cashflowCount + 1];
                object[,] expectedNominalValues_RiskNeutral = new object[2 * positions.Count + 1, cashflowCount + 1];
                expectedCashflowValues_RiskRente[0, 0] = "MarketValue";
                expectedCashflowValues_RiskNeutral[0, 0] = "MarketValue";
                expectedCouponsValues_RiskRente[0, 0] = "MarketValue";
                expectedCouponsValues_RiskNeutral[0, 0] = "MarketValue";
                expectedNominalValues_RiskRente[0, 0] = "MarketValue";
                expectedNominalValues_RiskNeutral[0, 0] = "MarketValue";
                for (int idxCashflow = 0; idxCashflow < cashflowCount; idxCashflow++)
                {
                    DateTime dateHeader = m_dStandardizedCashFlow_DatePoints[idxCashflow];
                    expectedCashflowValues_RiskRente[0, idxCashflow + 1] = dateHeader;
                    expectedCashflowValues_RiskNeutral[0, idxCashflow + 1] = dateHeader;
                    expectedCouponsValues_RiskRente[0, idxCashflow + 1] = dateHeader;
                    expectedCouponsValues_RiskNeutral[0, idxCashflow + 1] = dateHeader;
                    expectedNominalValues_RiskRente[0, idxCashflow + 1] = dateHeader;
                    expectedNominalValues_RiskNeutral[0, idxCashflow + 1] = dateHeader;

                    expectedCashflowValues_RiskRente_EUR[0, idxCashflow + 1] = dateHeader;
                    expectedCashflowValues_RiskNeutral_EUR[0, idxCashflow + 1] = dateHeader;
                    expectedCouponsValues_RiskRente_EUR[0, idxCashflow + 1] = dateHeader;
                    expectedCouponsValues_RiskNeutral_EUR[0, idxCashflow + 1] = dateHeader;
                    expectedNominalValues_RiskRente_EUR[0, idxCashflow + 1] = dateHeader;
                    expectedNominalValues_RiskNeutral_EUR[0, idxCashflow + 1] = dateHeader;
                }

                DateTime startTime = DateTime.Now;
                if (errors.CountErrors() == 0)
                {
                    for (int colDebug = 0; colDebug < debugValues_headers_List.Count; colDebug++)
                    {
                        debugValues[0, colDebug] = debugValues_headers_List.ElementAt(colDebug);
                        debugValues2[0, colDebug] = debugValues_headers_List.ElementAt(colDebug);
                    }

                    // Initialize progress bar
                    frm.InitProgBar(scenarios.Count);
                    string ccy, CCY;
                    Scenario scenario;
                    IndexCPI CPIindex;
                    // Calculate market value for all combinations of positions/scenarios
                    for (int idxScenario = 0; idxScenario < scenarios.Count; idxScenario++)
                    {
                        scenario = scenarios[idxScenario];
                        scenarioValuesLegFixed[0, idxScenario] = scenario.m_sName;
                        scenarioValuesLegFloat[0, idxScenario] = scenario.m_sName;
                        scenarioValues[0, idxScenario] = scenario.m_sName;
                        scenarioValuesBothLegs[0, idxScenario] = scenario.m_sName;
                        frm.SetStatus("Verwerking " + name + " scenario " + scenario.m_sName);

                        // Create scenario discount curve
                        CurveList scenarioZeroCurves = new CurveList();
                        SortedList<string, double> hullWhiteA = new SortedList<string, double>();
                        SortedList<string, double> hullWhiteSigma = new SortedList<string, double>();
                        foreach (ScenarioCurve scenarioCurve in scenario.m_YieldCurves)
                        {
                            ccy = scenarioCurve.m_sName;
                            CCY = ccy.ToUpper();
                            scenarioZeroCurves.Add(CCY, scenarioCurve.m_Curve);
                            hullWhiteA.Add(CCY, scenario.GetHullWhiteMeanReversion(ccy));
                            hullWhiteSigma.Add(CCY, scenario.GetHullWhiteVolatility(ccy));
                        }
                        CurveList scenarioEONIA_SpreadCurves = new CurveList();
                        foreach (ScenarioCurve scenarioCurve in scenario.m_EONIA_SpreadCurves)
                        {
                            ccy = scenarioCurve.m_sName;
                            scenarioEONIA_SpreadCurves.Add(ccy.ToUpper(), scenarioCurve.m_Curve);
                        }
                        CurveList scenarioZeroInflationCurves = new CurveList();
                        SortedList<string, IndexCPI> scenarioIndexCPI_List = new SortedList<string, IndexCPI>();
                        foreach (ScenarioCurve scenarioCurve in scenario.m_InflationCurves)
                        {
                            ccy = scenarioCurve.m_sName;
                            scenarioZeroInflationCurves.Add(ccy.ToUpper(), scenarioCurve.m_Curve);
                            CPIindex = new IndexCPI();
                            CPIindex.SetInflationInstance(frm.m_ReportDate.Value, 100, scenarioCurve.m_Curve, 100);
                            scenarioIndexCPI_List.Add(ccy.ToUpper(), CPIindex);
                        }
                        // Process all positions
                        int numberOfPositions = positions.Count;
                        for (int idxPosition = 0; idxPosition < numberOfPositions; idxPosition++)
                        {
                            positionIndex = idxPosition;
                            Position position = positions[idxPosition];
                            Instrument_Swap instrument = (Instrument_Swap)position.m_Instrument;
                            sSecurutyID = position.m_sSecurityID_LL;
                            // Create scenario discount curve
                            string ccyFixedLeg = Position.TranslateCurrency_Curve(scenarioZeroCurves, instrument.m_OriginalDataLeg[0].m_sCurrency);
                            string ccyFloatLeg = Position.TranslateCurrency_Curve(scenarioZeroCurves, instrument.m_OriginalDataLeg[1].m_sCurrency);
                            Curve[] scenarioZeroCurve2 = new Curve[2];
                            scenarioZeroCurve2[0] = scenarioZeroCurves[ccyFixedLeg];
                            scenarioZeroCurve2[1] = scenarioZeroCurves[ccyFloatLeg];
                            Curve[] EONIA_SpreadCurve2 = new Curve[2];
                            EONIA_SpreadCurve2[0] = scenarioEONIA_SpreadCurves[ccyFixedLeg];
                            EONIA_SpreadCurve2[1] = scenarioEONIA_SpreadCurves[ccyFloatLeg];
                            // Create scenarion CPI index
                            CPIindex = null;
                            if (scenarioIndexCPI_List.Count > 0)
                            {
                                if (scenarioIndexCPI_List.ContainsKey(ccyFloatLeg))
                                {
                                    CPIindex = scenarioIndexCPI_List[ccyFloatLeg];
                                }
                                else
                                {
                                    CPIindex = null;
                                }
                            }
                            // Create scenario FX rate
                            double[] fxLevel = new double[2];
                            ccy = Position.TranslateCurrency_Fx(scenario, instrument.m_OriginalDataLeg[0].m_sCurrency);
                            fxLevel[0] = scenario.m_Fx.ByName(ccy).m_fShockValue;
                            ccy = Position.TranslateCurrency_Fx(scenario, instrument.m_OriginalDataLeg[1].m_sCurrency);
                            fxLevel[1] = scenario.m_Fx.ByName(ccy).m_fShockValue;


                            double modelValue = instrument.getPrice(dtReport.Value, scenarioZeroCurve2, EONIA_SpreadCurve2, CPIindex, fxLevel);
                            double marketValueScenario = instrument.m_fMarketPrice + (modelValue - instrument.m_fModelPrice);
                            scenarioValues[idxPosition + 1, idxScenario] = marketValueScenario;

                            int legID = 0;
                            double modelValueLegFixed = instrument.getPriceLegFixed(dtReport.Value, scenarioZeroCurve2[legID], EONIA_SpreadCurve2[legID], CPIindex, fxLevel[legID], legID);
                            double marketValueScenarioLegFixed = instrument.m_fMarketPriceLeg[legID] + (modelValueLegFixed - instrument.m_fModelPriceLeg[legID]);
                            scenarioValuesLegFixed[idxPosition + 1, idxScenario] = marketValueScenarioLegFixed;
                            legID = 1;
                            double modelValueLegFloat = 0;
                            double marketValueScenarioLegFloat = 0;
                            if (instrument.m_Type == SwapType.IRS)
                            {
                                modelValueLegFloat = instrument.getPriceLegFloat(dtReport.Value, scenarioZeroCurve2[legID], EONIA_SpreadCurve2[legID], fxLevel[legID]);
                            }
                            else
                            {
                                modelValueLegFloat = instrument.getPriceLegFixed(dtReport.Value, scenarioZeroCurve2[legID], EONIA_SpreadCurve2[legID], CPIindex, fxLevel[legID], legID);
                            }
                            marketValueScenarioLegFloat = instrument.m_fMarketPriceLeg[legID] + (modelValueLegFloat - instrument.m_fModelPriceLeg[legID]);
                            scenarioValuesLegFloat[idxPosition + 1, idxScenario] = marketValueScenarioLegFloat;

                            scenarioValuesBothLegs[2 * idxPosition + 1, idxScenario] = marketValueScenarioLegFixed;
                            scenarioValuesBothLegs[2 * idxPosition + 2, idxScenario] = marketValueScenarioLegFloat;

                            if (scenario.isFairValueScenario())
                            {
                                legID = 0;
                                double modelValue_up = instrument.getPriceLegFixed(dtReport.Value, scenarioZeroCurve2[legID], EONIA_SpreadCurve2[legID], CPIindex, fxLevel[legID], instrument.m_fDiscountImpliedSpread[legID] - 0.0001, legID);
                                double modelValue_down = instrument.getPriceLegFixed(dtReport.Value, scenarioZeroCurve2[legID], EONIA_SpreadCurve2[legID], CPIindex, fxLevel[legID], instrument.m_fDiscountImpliedSpread[legID] + 0.0001, legID);
                                double spreadDurationLegFixed = (modelValue_up - modelValue_down) / modelValueLegFixed / 0.0002;
                                legID = 1;
                                if (instrument.m_Type == SwapType.IRS)
                                {
                                    modelValue_up = instrument.getPriceLegFloat(dtReport.Value, scenarioZeroCurve2[legID], EONIA_SpreadCurve2[legID], fxLevel[legID], instrument.m_fDiscountImpliedSpread[legID] - 0.0001, instrument.m_fBECashFlowImpliedSpread[legID]);
                                    modelValue_down = instrument.getPriceLegFloat(dtReport.Value, scenarioZeroCurve2[legID], EONIA_SpreadCurve2[legID], fxLevel[legID], instrument.m_fDiscountImpliedSpread[legID] + 0.0001, instrument.m_fBECashFlowImpliedSpread[legID]);
                                }
                                else
                                {
                                    modelValue_up = instrument.getPriceLegFixed(dtReport.Value, scenarioZeroCurve2[legID], EONIA_SpreadCurve2[legID], CPIindex, fxLevel[legID], instrument.m_fDiscountImpliedSpread[legID] - 0.0001, legID);
                                    modelValue_down = instrument.getPriceLegFixed(dtReport.Value, scenarioZeroCurve2[legID], EONIA_SpreadCurve2[legID], CPIindex, fxLevel[legID], instrument.m_fDiscountImpliedSpread[legID] + 0.0001, legID);
                                }
                                double spreadDurationLegFloat = (modelValue_up - modelValue_down) / modelValueLegFloat / 0.0002;
                                List<object> debugValues_Row1_List = new List<object>();
                                debugValues_Row1_List.Add("'" + frm.m_sReportingPeriod);
                                debugValues_Row1_List.Add("'" + position.m_sDATA_Source);
                                debugValues_Row1_List.Add("'" + ((Instrument_Swap)position.m_Instrument).m_OriginalDataLeg[0].m_dRow);
                                debugValues_Row1_List.Add("'" + position.m_sPositionId);
                                debugValues_Row1_List.Add("'" + position.m_sSecurityName_LL);
                                debugValues_Row1_List.Add(position.m_fVolume);
                                debugValues_Row1_List.Add(instrument.Name());
                                debugValues_Row1_List.Add("'" + instrument.m_PaymentType.ToString());
                                debugValues_Row1_List.Add("'" + instrument.m_Type.ToString());
                                debugValues_Row1_List.Add("'" + instrument.m_OriginalDataLeg[0].m_sCurrency);
                                debugValues_Row1_List.Add("'" + instrument.m_OriginalDataLeg[1].m_sCurrency);
                                debugValues_Row1_List.Add((null == instrument.m_StartDate) ? (object)"" : instrument.m_StartDate);
                                debugValues_Row1_List.Add(instrument.m_MaturityDate);
                                debugValues_Row1_List.Add(instrument.m_OriginalDataLeg[0].m_dLegID);
                                debugValues_Row1_List.Add(instrument.m_OriginalDataLeg[1].m_dLegID);
                                debugValues_Row1_List.Add("'" + instrument.m_OriginalDataLeg[0].m_sCouponType);
                                debugValues_Row1_List.Add("'" + instrument.m_OriginalDataLeg[1].m_sCouponType);
                                debugValues_Row1_List.Add(instrument.m_OriginalDataLeg[0].m_bFixedLeg);
                                debugValues_Row1_List.Add(instrument.m_OriginalDataLeg[1].m_bFixedLeg);
                                debugValues_Row1_List.Add(instrument.m_OriginalDataLeg[0].m_PaymentType);
                                debugValues_Row1_List.Add(instrument.m_OriginalDataLeg[1].m_PaymentType);
                                debugValues_Row1_List.Add(instrument.m_OriginalDataLeg[0].m_fNominal);
                                debugValues_Row1_List.Add(instrument.m_OriginalDataLeg[1].m_fNominal);
                                debugValues_Row1_List.Add(instrument.m_fRateLeg[0]);
                                debugValues_Row1_List.Add(instrument.m_fRateLeg[1]);
                                debugValues_Row1_List.Add(instrument.m_dFrequencyLeg[0]);
                                debugValues_Row1_List.Add(instrument.m_dFrequencyLeg[1]);
                                debugValues_Row1_List.Add(instrument.getMaturity(dtReport.Value));
                                debugValues_Row1_List.Add(instrument.getTenor());
                                debugValues_Row1_List.Add("Swap Model " + instrument.m_dModelType);
                                debugValues_Row1_List.Add(instrument.m_fMarketPrice);
                                debugValues_Row1_List.Add(instrument.m_fModelPrice);
                                debugValues_Row1_List.Add(instrument.m_fDuration);
                                debugValues_Row1_List.Add(instrument.m_fMarketPriceLeg[0]);
                                debugValues_Row1_List.Add(instrument.m_fModelPriceLeg[0]);
                                debugValues_Row1_List.Add(instrument.m_fDurationLeg[0]);
                                debugValues_Row1_List.Add(instrument.m_fMarketPriceLeg[1]);
                                debugValues_Row1_List.Add(instrument.m_fModelPriceLeg[1]);
                                debugValues_Row1_List.Add(instrument.m_fDurationLeg[1]);
                                debugValues_Row1_List.Add(instrument.m_fDiscountImpliedSpread[0]);
                                debugValues_Row1_List.Add(instrument.m_fDiscountImpliedSpread[1]);
                                debugValues_Row1_List.Add(instrument.m_fBECashFlowImpliedSpread[0]);
                                debugValues_Row1_List.Add(instrument.m_fBECashFlowImpliedSpread[1]);
                                debugValues_Row1_List.Add(instrument.m_OriginalDataLeg[0].m_fFxRate);
                                debugValues_Row1_List.Add(instrument.m_OriginalDataLeg[1].m_fFxRate);
                                debugValues_Row1_List.Add(scenarioZeroCurve2[0].GetRate(instrument.getMaturity(dtReport.Value)));
                                debugValues_Row1_List.Add(scenarioZeroCurve2[1].GetRate(instrument.getMaturity(dtReport.Value)));
                                debugValues_Row1_List.Add(spreadDurationLegFixed);
                                debugValues_Row1_List.Add(spreadDurationLegFloat);
                                debugValues_Row1_List.Add(instrument.m_dIterations);
                                List<object> debugValues_Row2_List = new List<object>();
                                for (int colDebug = 0; colDebug < debugValues_headers_List.Count; colDebug++)
                                {
                                    if (2 == colDebug)
                                    {
                                        debugValues_Row2_List.Add("'" + ((Instrument_Swap)position.m_Instrument).m_OriginalDataLeg[1].m_dRow);
                                    }
                                    else
                                    {
                                        debugValues_Row2_List.Add(debugValues_Row1_List.ElementAt(colDebug));
                                    }
                                    debugValues[idxPosition + 1, colDebug] = debugValues_Row1_List.ElementAt(colDebug);
                                    debugValues2[2 * idxPosition + 1, colDebug] = debugValues_Row1_List.ElementAt(colDebug);
                                    debugValues2[2 * idxPosition + 2, colDebug] = debugValues_Row2_List.ElementAt(colDebug);
                                }
                                position.m_fFairValue = marketValueScenario;
                                if (isCashFlowNeeded)
                                {
                                    Curve scenarioZeroCurveEUR = scenarioZeroCurves["EUR"];
                                    instrument.m_CashFlow_FXRate_Reported = new double[2][][];
                                    instrument.m_CashFlow_Reported = new double[2][][];
                                    instrument.m_CashFlow_Coupons_Reported = new double[2][][];
                                    instrument.m_CashFlow_Nominal_Reported = new double[2][][];
                                    for (int i = 0; i < 2; i++)
                                    {
                                        instrument.m_CashFlow_FXRate_Reported[i] = new double[2][];
                                        instrument.m_CashFlow_Reported[i] = new double[2][];
                                        instrument.m_CashFlow_Coupons_Reported[i] = new double[2][];
                                        instrument.m_CashFlow_Nominal_Reported[i] = new double[2][];
                                        instrument.calculateStandardizedCashFlows(dtReport.Value, scenarioZeroCurve2, scenarioZeroCurveEUR, m_dStandardizedCashFlow_DatePoints, i, CashFlowType.RiskNeutral);
                                        instrument.calculateStandardizedCashFlows(dtReport.Value, scenarioZeroCurve2, scenarioZeroCurveEUR, m_dStandardizedCashFlow_DatePoints, i, CashFlowType.RiskRente);
                                    }
                                    int row_1 = 2 * idxPosition + 1;
                                    int row_2 = 2 * idxPosition + 2;
                                    // CF:
                                    expectedCashflowValues_RiskRente[row_1, 0] = scenarioValuesBothLegs[row_1, idxScenario];
                                    expectedCashflowValues_RiskRente[row_2, 0] = scenarioValuesBothLegs[row_2, idxScenario];
                                    expectedCashflowValues_RiskNeutral[row_1, 0] = scenarioValuesBothLegs[row_1, idxScenario];
                                    expectedCashflowValues_RiskNeutral[row_2, 0] = scenarioValuesBothLegs[row_2, idxScenario];
                                    expectedCashflowValues_RiskRente_EUR[row_1, 0] = scenarioValuesBothLegs[row_1, idxScenario];
                                    expectedCashflowValues_RiskRente_EUR[row_2, 0] = scenarioValuesBothLegs[row_2, idxScenario];
                                    expectedCashflowValues_RiskNeutral_EUR[row_1, 0] = scenarioValuesBothLegs[row_1, idxScenario];
                                    expectedCashflowValues_RiskNeutral_EUR[row_2, 0] = scenarioValuesBothLegs[row_2, idxScenario];
                                    // Coupons:
                                    expectedCouponsValues_RiskRente[row_1, 0] = scenarioValuesBothLegs[row_1, idxScenario];
                                    expectedCouponsValues_RiskRente[row_2, 0] = scenarioValuesBothLegs[row_2, idxScenario];
                                    expectedCouponsValues_RiskNeutral[row_1, 0] = scenarioValuesBothLegs[row_1, idxScenario];
                                    expectedCouponsValues_RiskNeutral[row_2, 0] = scenarioValuesBothLegs[row_2, idxScenario];
                                    expectedCouponsValues_RiskRente_EUR[row_1, 0] = scenarioValuesBothLegs[row_1, idxScenario];
                                    expectedCouponsValues_RiskRente_EUR[row_2, 0] = scenarioValuesBothLegs[row_2, idxScenario];
                                    expectedCouponsValues_RiskNeutral_EUR[row_1, 0] = scenarioValuesBothLegs[row_1, idxScenario];
                                    expectedCouponsValues_RiskNeutral_EUR[row_2, 0] = scenarioValuesBothLegs[row_2, idxScenario];
                                    // Nominal:
                                    expectedNominalValues_RiskRente[row_1, 0] = scenarioValuesBothLegs[row_1, idxScenario];
                                    expectedNominalValues_RiskRente[row_2, 0] = scenarioValuesBothLegs[row_2, idxScenario];
                                    expectedNominalValues_RiskNeutral[row_1, 0] = scenarioValuesBothLegs[row_1, idxScenario];
                                    expectedNominalValues_RiskNeutral[row_2, 0] = scenarioValuesBothLegs[row_2, idxScenario];
                                    expectedNominalValues_RiskRente_EUR[row_1, 0] = scenarioValuesBothLegs[row_1, idxScenario];
                                    expectedNominalValues_RiskRente_EUR[row_2, 0] = scenarioValuesBothLegs[row_2, idxScenario];
                                    expectedNominalValues_RiskNeutral_EUR[row_1, 0] = scenarioValuesBothLegs[row_1, idxScenario];
                                    expectedNominalValues_RiskNeutral_EUR[row_2, 0] = scenarioValuesBothLegs[row_2, idxScenario];
                                    // FX:
                                    double[] FX_RiskRente_0 = instrument.m_CashFlow_FXRate_Reported[0][(int)CashFlowType.RiskRente];
                                    double[] FX_RiskNeutral_0 = instrument.m_CashFlow_FXRate_Reported[0][(int)CashFlowType.RiskNeutral];
                                    double[] FX_RiskRente_1 = instrument.m_CashFlow_FXRate_Reported[1][(int)CashFlowType.RiskRente];
                                    double[] FX_RiskNeutral_1 = instrument.m_CashFlow_FXRate_Reported[1][(int)CashFlowType.RiskNeutral];
                                    // CF:
                                    double[] CF_RiskRente_0 = instrument.m_CashFlow_Reported[0][(int)CashFlowType.RiskRente];
                                    double[] CF_RiskNeutral_0 = instrument.m_CashFlow_Reported[0][(int)CashFlowType.RiskNeutral];
                                    double[] CF_RiskRente_1 = instrument.m_CashFlow_Reported[1][(int)CashFlowType.RiskRente];
                                    double[] CF_RiskNeutral_1 = instrument.m_CashFlow_Reported[1][(int)CashFlowType.RiskNeutral];
                                    // Coupons:
                                    double[] Coupons_RiskRente_0 = instrument.m_CashFlow_Coupons_Reported[0][(int)CashFlowType.RiskRente];
                                    double[] Coupons_RiskNeutral_0 = instrument.m_CashFlow_Coupons_Reported[0][(int)CashFlowType.RiskNeutral];
                                    double[] Coupons_RiskRente_1 = instrument.m_CashFlow_Coupons_Reported[1][(int)CashFlowType.RiskRente];
                                    double[] Coupons_RiskNeutral_1 = instrument.m_CashFlow_Coupons_Reported[1][(int)CashFlowType.RiskNeutral];
                                    // Nominal:
                                    double[] Nominal_RiskRente_0 = instrument.m_CashFlow_Nominal_Reported[0][(int)CashFlowType.RiskRente];
                                    double[] Nominal_RiskNeutral_0 = instrument.m_CashFlow_Nominal_Reported[0][(int)CashFlowType.RiskNeutral];
                                    double[] Nominal_RiskRente_1 = instrument.m_CashFlow_Nominal_Reported[1][(int)CashFlowType.RiskRente];
                                    double[] Nominal_RiskNeutral_1 = instrument.m_CashFlow_Nominal_Reported[1][(int)CashFlowType.RiskNeutral];
                                    for (int idxCashflow = 0; idxCashflow < cashflowCount; idxCashflow++)
                                    {
                                        // CF:
                                        expectedCashflowValues_RiskRente[row_1, idxCashflow + 1] = CF_RiskRente_0[idxCashflow] * instrument.m_fNominalLeg[0];
                                        expectedCashflowValues_RiskRente[row_2, idxCashflow + 1] = CF_RiskRente_1[idxCashflow] * instrument.m_fNominalLeg[1];
                                        expectedCashflowValues_RiskNeutral[row_1, idxCashflow + 1] = CF_RiskNeutral_0[idxCashflow] * instrument.m_fNominalLeg[0];
                                        expectedCashflowValues_RiskNeutral[row_2, idxCashflow + 1] = CF_RiskNeutral_1[idxCashflow] * instrument.m_fNominalLeg[1];
                                        expectedCashflowValues_RiskRente_EUR[row_1, idxCashflow + 1] = CF_RiskRente_0[idxCashflow] * instrument.m_fNominalLeg[0] * FX_RiskRente_0[idxCashflow];
                                        expectedCashflowValues_RiskRente_EUR[row_2, idxCashflow + 1] = CF_RiskRente_1[idxCashflow] * instrument.m_fNominalLeg[1] * FX_RiskRente_1[idxCashflow];
                                        expectedCashflowValues_RiskNeutral_EUR[row_1, idxCashflow + 1] = CF_RiskNeutral_0[idxCashflow] * instrument.m_fNominalLeg[0] * FX_RiskNeutral_0[idxCashflow];
                                        expectedCashflowValues_RiskNeutral_EUR[row_2, idxCashflow + 1] = CF_RiskNeutral_1[idxCashflow] * instrument.m_fNominalLeg[1] * FX_RiskNeutral_1[idxCashflow];
                                        // Coupons:
                                        expectedCouponsValues_RiskRente[row_1, idxCashflow + 1] = Coupons_RiskRente_0[idxCashflow] * instrument.m_fNominalLeg[0];
                                        expectedCouponsValues_RiskRente[row_2, idxCashflow + 1] = Coupons_RiskRente_1[idxCashflow] * instrument.m_fNominalLeg[1];
                                        expectedCouponsValues_RiskNeutral[row_1, idxCashflow + 1] = Coupons_RiskNeutral_0[idxCashflow] * instrument.m_fNominalLeg[0];
                                        expectedCouponsValues_RiskNeutral[row_2, idxCashflow + 1] = Coupons_RiskNeutral_1[idxCashflow] * instrument.m_fNominalLeg[1];
                                        expectedCouponsValues_RiskRente_EUR[row_1, idxCashflow + 1] = Coupons_RiskRente_0[idxCashflow] * instrument.m_fNominalLeg[0] * FX_RiskRente_0[idxCashflow];
                                        expectedCouponsValues_RiskRente_EUR[row_2, idxCashflow + 1] = Coupons_RiskRente_1[idxCashflow] * instrument.m_fNominalLeg[1] * FX_RiskRente_1[idxCashflow];
                                        expectedCouponsValues_RiskNeutral_EUR[row_1, idxCashflow + 1] = Coupons_RiskNeutral_0[idxCashflow] * instrument.m_fNominalLeg[0] * FX_RiskNeutral_0[idxCashflow];
                                        expectedCouponsValues_RiskNeutral_EUR[row_2, idxCashflow + 1] = Coupons_RiskNeutral_1[idxCashflow] * instrument.m_fNominalLeg[1] * FX_RiskNeutral_1[idxCashflow];
                                        // Nominal:
                                        expectedNominalValues_RiskRente[row_1, idxCashflow + 1] = Nominal_RiskRente_0[idxCashflow] * instrument.m_fNominalLeg[0];
                                        expectedNominalValues_RiskRente[row_2, idxCashflow + 1] = Nominal_RiskRente_1[idxCashflow] * instrument.m_fNominalLeg[1];
                                        expectedNominalValues_RiskNeutral[row_1, idxCashflow + 1] = Nominal_RiskNeutral_0[idxCashflow] * instrument.m_fNominalLeg[0];
                                        expectedNominalValues_RiskNeutral[row_2, idxCashflow + 1] = Nominal_RiskNeutral_1[idxCashflow] * instrument.m_fNominalLeg[1];
                                        expectedNominalValues_RiskRente_EUR[row_1, idxCashflow + 1] = Nominal_RiskRente_0[idxCashflow] * instrument.m_fNominalLeg[0] * FX_RiskRente_0[idxCashflow];
                                        expectedNominalValues_RiskRente_EUR[row_2, idxCashflow + 1] = Nominal_RiskRente_1[idxCashflow] * instrument.m_fNominalLeg[1] * FX_RiskRente_1[idxCashflow];
                                        expectedNominalValues_RiskNeutral_EUR[row_1, idxCashflow + 1] = Nominal_RiskNeutral_0[idxCashflow] * instrument.m_fNominalLeg[0] * FX_RiskNeutral_0[idxCashflow];
                                        expectedNominalValues_RiskNeutral_EUR[row_2, idxCashflow + 1] = Nominal_RiskNeutral_1[idxCashflow] * instrument.m_fNominalLeg[1] * FX_RiskNeutral_0[idxCashflow];
                                    }
                                }

                            }

                            System.Windows.Forms.Application.DoEvents();
                        }

                        // Update progress bar
                        frm.IncrementProgBar();
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                DateTime endTime = DateTime.Now;
                double runTime = (endTime - startTime).TotalSeconds;

                // Write aggregated market values to results workbook
                frm.SetStatus("Opslaan resultaten " + name);
                object[,] curveValues;
                Scenario baseScenario = scenarios.getScenarioFairValue();
                CurveList baseCurves = new CurveList();
                foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
                {
                    baseCurves.Add(scenarioCurve.m_sName.ToUpper(), scenarioCurve.m_Curve);
                }
                foreach (ScenarioCurve scenarioCurve in baseScenario.m_EONIA_SpreadCurves)
                {
                    baseCurves.Add("EONIA_" + scenarioCurve.m_sName.ToUpper(), scenarioCurve.m_Curve);
                }
                curveValues = baseCurves.ToArray(true);
                object[,] errorValues = errors.ToArray();
//                frm.WriteToExcel(name, positions, scenarioValues, curveValues, debugValues, errorValues, runTime, null, null);
                if (errors.CountErrors() > 0)
                {
                    MessageBox.Show("Fouten in verwerking " + name + ", bekijk Error werkblad in output bestand", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
//                frm.WriteToExcel(name + " Fixed leg", positions, scenarioValuesLegFixed, curveValues, debugValues, errorValues, runTime, null, null, false);
//                frm.WriteToExcel(name + " Float leg", positions, scenarioValuesLegFloat, curveValues, debugValues, errorValues, runTime, null, null, false);
                frm.WriteSwapLegsToExcel(name + " Both legs", positions, scenarioValuesBothLegs, curveValues, debugValues2, errorValues, runTime, null, null);
                if (isCashFlowNeeded)
                {
                    // CF:
                    frm.WriteSwapLegsToExcel(name + " Cash flows Risk Neutral", positions, expectedCashflowValues_RiskNeutral, curveValues, debugValues, errorValues, runTime, null, null);
                    frm.WriteSwapLegsToExcel(name + " Cash flows Risk Rente", positions, expectedCashflowValues_RiskRente, curveValues, debugValues, errorValues, runTime, null, null);
                    frm.WriteSwapLegsToExcel(name + " Cash flows Risk Neutral (EUR)", positions, expectedCashflowValues_RiskNeutral_EUR, curveValues, debugValues, errorValues, runTime, null, null);
                    frm.WriteSwapLegsToExcel(name + " Cash flows Risk Rente (EUR)", positions, expectedCashflowValues_RiskRente_EUR, curveValues, debugValues, errorValues, runTime, null, null);
                    // Coupons:
                    frm.WriteSwapLegsToExcel(name + " Coupons Risk Neutral", positions, expectedCouponsValues_RiskNeutral, curveValues, debugValues, errorValues, runTime, null, null);
                    frm.WriteSwapLegsToExcel(name + " Coupons Risk Rente", positions, expectedCouponsValues_RiskRente, curveValues, debugValues, errorValues, runTime, null, null);
                    frm.WriteSwapLegsToExcel(name + " Coupons Risk Neutral (EUR)", positions, expectedCouponsValues_RiskNeutral_EUR, curveValues, debugValues, errorValues, runTime, null, null);
                    frm.WriteSwapLegsToExcel(name + " Coupons Risk Rente (EUR)", positions, expectedCouponsValues_RiskRente_EUR, curveValues, debugValues, errorValues, runTime, null, null);
                    // Nominal:
                    frm.WriteSwapLegsToExcel(name + " Nominal Risk Neutral", positions, expectedNominalValues_RiskNeutral, curveValues, debugValues, errorValues, runTime, null, null);
                    frm.WriteSwapLegsToExcel(name + " Nominal Risk Rente", positions, expectedNominalValues_RiskRente, curveValues, debugValues, errorValues, runTime, null, null);
                    frm.WriteSwapLegsToExcel(name + " Nominal Risk Neutral (EUR)", positions, expectedNominalValues_RiskNeutral_EUR, curveValues, debugValues, errorValues, runTime, null, null);
                    frm.WriteSwapLegsToExcel(name + " Nominal Risk Rente (EUR)", positions, expectedNominalValues_RiskRente_EUR, curveValues, debugValues, errorValues, runTime, null, null);
                }
            }
            catch (Exception exc)
            {
                string errorMessage = "Fout in verwerking " + name + " SecurityID_LL (" + sSecurutyID + ") :\n" + exc.Message;
                MessageBox.Show(errorMessage, "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            frm.HideProgBar();
        }
        public void Process_Swaptions(FormMain frm, string name, DateTime dtReport, PositionList positions, 
            ScenarioList scenarios, ErrorList errors)
        {
            try
            {
                List<string> headerNames = new List<string>();
                headerNames.Add("Period");
                headerNames.Add("Report Date");
                headerNames.Add("DATA SOURCE");
                headerNames.Add("Row ID Source");
                headerNames.Add("Row ID Debug");
                headerNames.Add("SelectieIndex_LL");
                headerNames.Add("Is Look-Through Data");
                headerNames.Add("PositionId");
                headerNames.Add("SecurityName");
                headerNames.Add("Volume");
                headerNames.Add("InstrumentType");
                headerNames.Add("SwaptionType");
                headerNames.Add("SettlementType");
                headerNames.Add("ExpiryDate");
                headerNames.Add("MaturityDate");
                headerNames.Add("StrikeRate");
                headerNames.Add("Expiry");
                headerNames.Add("Tenor");
                headerNames.Add("MarketPrice");
                headerNames.Add("TheoreticalPrice");
                headerNames.Add("Theoretical1%Vol");
                headerNames.Add("IntrisicValue");
                headerNames.Add("ForwardRate");
                headerNames.Add("Volatility");
                headerNames.Add("ImpliedVol");
                headerNames.Add("ATM_Vol");
                headerNames.Add("Iterations");
                headerNames.Add("Discount rate");
                headerNames.Add("Comments on Moneyness");
                object[,] scenarioValues = new object[positions.Count + 1, scenarios.Count];
                object[,] debugValues = new object[positions.Count + 1, headerNames.Count];
                Scenario baseScenario;
                DateTime startTime = DateTime.Now;
                if (errors.CountErrors() == 0)
                {
                    int colDebug = 0;
                    for (colDebug = 0; colDebug < headerNames.Count; colDebug++)
                    {
                        debugValues[0, colDebug] = "'" + headerNames[colDebug];
                    }
                    // Initialize progress bar
                    frm.InitProgBar(scenarios.Count);
                    Curve scenarioZeroCurve;
                    string ccy;
                    Scenario scenario;
                    // Calculate market value for all combinations of positions/scenarios
                    for (int idxScenario = 0; idxScenario < scenarios.Count; idxScenario++)
                    {
                        scenario = scenarios[idxScenario];
                        scenarioValues[0, idxScenario] = scenario.m_sName;
                        frm.SetStatus("Verwerking " + name + " scenario " + scenario.m_sName);

                        // Create scenario discount curve
                        CurveList scenarioZeroCurves = new CurveList();
                        SortedList<string, double> hullWhiteA = new SortedList<string, double>();
                        SortedList<string, double> hullWhiteSigma = new SortedList<string, double>();
                        foreach (ScenarioCurve scenarioCurve in scenario.m_YieldCurves)
                        {
                            ccy = scenarioCurve.m_sName;
                            scenarioZeroCurve = scenarioCurve.m_Curve;
                            scenarioZeroCurves.Add(ccy.ToUpper(), scenarioZeroCurve);
                            hullWhiteA.Add(ccy.ToUpper(), scenario.GetHullWhiteMeanReversion(ccy));
                            hullWhiteSigma.Add(ccy.ToUpper(), scenario.GetHullWhiteVolatility(ccy));
                        }
                        CurveList scenarioEONIA_SpreadCurves = new CurveList();
                        foreach (ScenarioCurve scenarioCurve in scenario.m_EONIA_SpreadCurves)
                        {
                            ccy = scenarioCurve.m_sName;
                            scenarioEONIA_SpreadCurves.Add(ccy.ToUpper(), scenarioCurve.m_Curve);
                        }
                        // Process all positions
                        for (int idxPosition = 0; idxPosition < positions.Count; idxPosition++)
                        {
                            int row = idxPosition + 1;
                            Position position = positions[idxPosition];
                            Instrument_Swaption instrument = (Instrument_Swaption)position.m_Instrument;

                            // Create scenario discount curve
                            ccy = Position.TranslateCurrency_Curve(scenarioZeroCurves, position.m_sCurrency);
                            scenarioZeroCurve = scenarioZeroCurves[ccy];
                            Curve EONIA_SpreadCurve = scenarioEONIA_SpreadCurves[ccy];
                            Curve zeroEONIACurve = scenarioZeroCurve + EONIA_SpreadCurve;
                            double modelValue = 0;
                            double marketValueScenario = 0;
                            double forwardRate = instrument.getForwardRate(dtReport, scenarioZeroCurve);
                            double swaptionVolatility = instrument.m_fVolatility;
                            if (scenario.ExistsVolatilityShock(ccy))
                            {
                                ScenarioValue scenarioValue = scenario.GetVolatilityShock(ccy);
                                swaptionVolatility = scenarioValue.ShockedValue(instrument.m_fVolatility);
                            }
                            if (m_HullWhitModel)
                            {
                                double meanReversion = hullWhiteA[ccy];
                                double sigma = hullWhiteSigma[ccy];
                                modelValue = instrument.getPrice(dtReport, scenarioZeroCurve, meanReversion, sigma);
                            }
                            else
                            {
                                modelValue = instrument.getPrice(dtReport, scenarioZeroCurve, swaptionVolatility);
//                                modelValue = instrument.getPrice(dtReport, scenarioZeroCurve, zeroEONIACurve, swaptionVolatility);
                            }
                            marketValueScenario = position.m_fVolume * Math.Max(0, instrument.m_fMarketPrice + (modelValue - instrument.m_fModelPrice));

                            scenarioValues[idxPosition + 1, idxScenario] = marketValueScenario;
                            if (scenario.isFairValueScenario())
                            {
                                colDebug = 0;
                                debugValues[row, colDebug++] = "'" + frm.m_sReportingPeriod;
                                debugValues[row, colDebug++] = dtReport;
                                debugValues[row, colDebug++] = "'" + position.m_sDATA_Source;
                                debugValues[row, colDebug++] = "'" + position.m_sRow;
                                debugValues[row, colDebug++] = "'" + (row + 1).ToString();
                                debugValues[row, colDebug++] = "'" + position.m_sSelectieIndex_LL;
                                debugValues[row, colDebug++] = "'" + position.m_bIsLookThroughPosition;
                                debugValues[row, colDebug++] = "'" + position.m_sPositionId;
                                debugValues[row, colDebug++] = "'" + position.m_sSecurityName_LL;
                                debugValues[row, colDebug++] = position.m_fVolume;
                                debugValues[row, colDebug++] = "'" + instrument.Name();
                                debugValues[row, colDebug++] = "'" + instrument.m_Type.ToString();
                                debugValues[row, colDebug++] = (instrument.m_bCashSettled) ? "cash" : "physical";
                                debugValues[row, colDebug++] = instrument.m_ExpiryDate;
                                debugValues[row, colDebug++] = instrument.m_MaturityDate;
                                debugValues[row, colDebug++] = instrument.m_fStrike;
                                debugValues[row, colDebug++] = instrument.getMaturity(dtReport);
                                debugValues[row, colDebug++] = instrument.getTenor();
                                debugValues[row, colDebug++] = instrument.m_fMarketPrice;
                                debugValues[row, colDebug++] = instrument.m_fModelPrice;
                                debugValues[row, colDebug++] = instrument.m_fCleanPriceOnePercentVol;
                                debugValues[row, colDebug++] = instrument.m_fIntrisicValue;
                                debugValues[row, colDebug++] = instrument.m_fForwardRate;
                                debugValues[row, colDebug++] = instrument.m_fVolatility;
                                if (double.IsInfinity(instrument.m_fImpliedVol))
                                {
                                    debugValues[row, colDebug++] = "#N/A";
                                }
                                else
                                {
                                    debugValues[row, colDebug++] = instrument.m_fImpliedVol;
                                }
                                debugValues[row, colDebug++] = "N.A."; // ATM vol
                                debugValues[row, colDebug++] = instrument.m_dIterations;
                                debugValues[row, colDebug++] = scenarioZeroCurve.GetRate(instrument.getMaturity(dtReport));
                                debugValues[row, colDebug++] = "'" + instrument.m_sCommentOnMoneyness;

                                position.m_fFairValue = marketValueScenario;
                            }

                            System.Windows.Forms.Application.DoEvents();
                        }

                        // Update progress bar
                        frm.IncrementProgBar();
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
                DateTime endTime = DateTime.Now;
                double runTime = (endTime - startTime).TotalSeconds;

                // Write aggregated market values to results workbook
                frm.SetStatus("Opslaan resultaten " + name);
                baseScenario = scenarios.getScenarioFairValue();
                CurveList baseCurves = new CurveList();
                foreach (ScenarioCurve scenarioCurve in baseScenario.m_YieldCurves)
                {
                    baseCurves.Add(scenarioCurve.m_sName.ToUpper(), scenarioCurve.m_Curve);
                }
                foreach (ScenarioCurve scenarioCurve in baseScenario.m_EONIA_SpreadCurves)
                {
                    baseCurves.Add("EONIA_" + scenarioCurve.m_sName.ToUpper(), scenarioCurve.m_Curve);
                }
                object[,] curveValues = baseCurves.ToArray(true);
                object[,] errorValues = errors.ToArray();
                frm.WriteToExcel(name, positions, scenarioValues, curveValues, debugValues, errorValues, runTime, null, null, false);

                if (errors.CountErrors() > 0)
                    MessageBox.Show("Fouten in verwerking " + name + ", bekijk Error werkblad in output bestand", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (IOException exc)
            {
                MessageBox.Show("Fout in verwerking " + name + ":\n" + exc.Message, "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            frm.HideProgBar();
        }
        public DateTime[] getReportedCashFlowDates(DateTime? reportDate, string cashflowsBucketType, int cashflowsMaxBucket)
        {
            DateTime? FirstDate_Model = reportDate.EndOfMonth().Value.AddDays(+1);
            DateTime[] StandardizedCashFlow_DatePoints; 
            if (cashflowsBucketType.Substring(0, 1).Equals("M", StringComparison.InvariantCultureIgnoreCase))
            {
                StandardizedCashFlow_DatePoints = new DateTime[12 * cashflowsMaxBucket + 1];
                StandardizedCashFlow_DatePoints[0] = FirstDate_Model.Value;
                for (int i = 0; i < cashflowsMaxBucket; i++)
                {
                    for (int j = 0; j < 12; j++)
                    {
                        int index = 12 * i + j;
                        StandardizedCashFlow_DatePoints[index + 1] = StandardizedCashFlow_DatePoints[index].AddMonths(1);
                    }
                }
            }
            else
            {
                StandardizedCashFlow_DatePoints = new DateTime[cashflowsMaxBucket + 1];
                StandardizedCashFlow_DatePoints[0] = FirstDate_Model.Value;
                for (int i = 0; i < cashflowsMaxBucket; i++)
                {
                    StandardizedCashFlow_DatePoints[i + 1] = StandardizedCashFlow_DatePoints[i].AddYears(1);
                }
            }
            return StandardizedCashFlow_DatePoints;
        }
    }

}
