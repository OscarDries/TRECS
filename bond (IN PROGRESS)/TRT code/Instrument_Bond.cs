using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TotalRisk.ExcelWrapper;
using TotalRisk.Utilities;

namespace TotalRisk.ValuationModule
{
    public enum BondType
    {
        CASH_FLOW = (int)0,
        FIXED = (int)1,
        FLOATER = (int)2,
        INFLATION = (int) 3
    }
    public enum DNB_Bond
    {
        Undefined = (int)0,
        Government = (int)1,
        Corporate = (int)2,
        Covered = (int)3
    }
    public class Instrument_Bond_OriginalData
    {
        public DateTime? m_dtReport;
        public int m_dRow = -1;
        public string m_sUniquePositionId;
        public string m_sSelectieIndex_LL;
        public bool m_bLookthroughData;
        public BondType m_BondType;
        public DNB_Bond m_DNB_Type = DNB_Bond.Undefined;
        public string m_sDNB_CountryUnion = "";
        public bool m_bDNB_Financial = false;
        public string m_sDNB_Rating = "";
        public string m_sScope3 = "";
        public bool m_bICO;
        public string m_sPortfolioID = "";
        public string m_sSecurityID = "";
        public string m_sSecurityName = "";
        public string m_sSecurityID_LL = "";
        public string m_sSecurityName_LL = "";
        public string m_sSecurityType_LL = "";
        public string m_sLeg = "";
        public string m_sType = "";
        public string m_sCurrencyCountry = "";
        public string m_sCurrency = "";
        public double m_fExpiryDate;
        public double m_fCallDate;
        public double m_fCoupon; // %
        public string m_sCouponType;
        public int m_dCouponFrequency;
        public double m_fCouponReferenceRate;
        public double m_fCouponSpread; // %
        public DateTime? m_dtFirstCouponDate;
        public double m_fFxRate;
        public double m_fNominal;
        public double m_fMarketValue;
        public double m_fAccruedInterest_LL;
        public string m_sCIC_ID = "";
        public string m_sCIC_ID_LL = "";
        public bool m_bGovGuarantee = false;
        public bool m_bEEA;
        public int m_dSecuritisationType;
        public double m_fCollateral;
        public string m_sIssuerCreditQuality;
        public int m_dIssuerCreditQuality;
        public string m_sSecurityCreditQuality;
        public int m_dSecurityCreditQuality;
        public string m_sGroupCounterpartyName;
        public string m_sGroupCounterpartyLEI;
        public string m_sGroupCounterpartyCQS;
        public int m_dGroupCounterpartyCQS;
        public bool m_bCallable;
        public double m_fModifiedDurationOrig;
        public double m_fModifiedDurationCorrected;
        public double m_fSpreadDurationOrig;
        public double m_fSpreadDurationCorrected;

        public string m_sNACEcode = "";
        public string m_sAccount_LL = "";
        public string m_sAccount = "";
        public string m_sECAP_Category_LL = "";
        public string m_sPortfolioPurpose = "";
        public string m_sDATA_Source = "";

    }
    public class Instrument_Bond_OriginalData_List : List<Instrument_Bond_OriginalData>
    {
        public Instrument_Bond_OriginalData_List() { }
        public Instrument_Bond_OriginalData_List(Instrument_Bond_OriginalData_List pList)
            : this()
        {
            foreach (Instrument_Bond_OriginalData Position in pList)
            {
                Instrument_Bond_OriginalData p = Position;
                this.Add(p);
            }
        }

    }

    public class CBondDebugData
    {
        public string m_sReportingPeriod;
        public DateTime? m_ReportDate;
        public int m_dNumberOfLinks = 0; // the number of linkis to the next period. It is filled in when analyzed the next period
        public string m_sRow_Source = "-1"; // the row id in the original data file
        public string m_sRow_Debug = "-1"; // the row ID in the debug tab of Total Risk Output
        public string m_sRow_Debug_Linked = "-1"; // the first linked position of the previous period
        public string m_sScope3;
        public string m_sPosition_ID;
        public bool m_bIsLookThroughData;
        public string m_sPortfolio_ID;
        public string m_sSecurity_ID;
        public string m_sSecurity_ID_LL;
        public string m_sSecurity_Type;
        public string m_sCIC;
        public string m_sCIC_LL;
        public double m_fNominal_Value;
        public double m_fMarket_Value;
        public double m_fImplied_Spread;
        public double m_fFxRate;

        public CBondDebugData()
        {
        }

        public CBondDebugData(CBondDebugData p)
            : this()
        {
            m_sReportingPeriod = p.m_sReportingPeriod;
            m_ReportDate = p.m_ReportDate;
            m_dNumberOfLinks = p.m_dNumberOfLinks;
            m_sRow_Source = p.m_sRow_Source;
            m_sRow_Debug = p.m_sRow_Debug;
            m_sScope3 = p.m_sScope3;
            m_sPosition_ID = p.m_sPosition_ID;
            m_bIsLookThroughData = p.m_bIsLookThroughData;
            m_sPortfolio_ID = p.m_sPortfolio_ID;
            m_sSecurity_ID = p.m_sSecurity_ID;
            m_sSecurity_ID_LL = p.m_sSecurity_ID_LL;
            m_sSecurity_Type = p.m_sSecurity_Type;
            m_sCIC = p.m_sCIC;
            m_sCIC_LL = p.m_sCIC_LL;
            m_fNominal_Value = p.m_fNominal_Value;
            m_fMarket_Value = p.m_fMarket_Value;
            m_fImplied_Spread = p.m_fImplied_Spread;
            m_fFxRate = p.m_fFxRate;
        }
    }
    public class CBondDebugDataList : List<CBondDebugData>
    {
        //        public Dictionary<string, int> headers;
        public CBondDebugDataList() { }

        public CBondDebugDataList(CBondDebugDataList pList)
            : this()
        {
            foreach (CBondDebugData Position in pList)
            {
                CBondDebugData p = Position;
                this.Add(p);
            }
        }
        // Methods
        public int Add(CBondDebugDataList list)
        {
            foreach (CBondDebugData entry in list)
            {
                Add(entry);
            }
            return this.Count;
        }
    }


    /// <summary>
    /// This class define a swaption instrument
    /// </summary>
    public class Instrument_Bond : FinancialInstrument
    {// 
        /****  MODEL VERSION NUMBER   ******/
        static public string getVersion()
        {
            string version = "Version 5.04: created on 10.06.2022: inflation bond is introduced."; //    
            return version;
        }

        /**** DATA  ***/
        public Instrument_Bond_OriginalData m_OriginalBond;
        public Import_ValuationModule.CFixedIncomeCashFlowData[] m_OriginalCashFlowData;
        public CBondDebugData m_BondDebugData_PrevPeriod = null;
        public CBondDebugData m_BondDebugData_CurrentPeriod = null;
        // Properties
        public BondType m_Type;
        public bool m_bLookthroughBond = false;
        public DateTime m_MaturityDate { get; set; } /// Expiry date for swap
        public DateTime m_CallDate { get; set; } /// Call date for swap
        public bool m_bCallable = false;
        public bool m_bDefaulted = false;
        public double m_fNominal { get; set; }
        public double m_fMarketPrice { get; set; } // Market price in EUR
        public double m_fModelPrice { get; set; } // Theoretical price in EUR
        public double m_fModelValueAtZeroSpread { get; set; } // Theoretic value with the zero spread  
        public double m_fDuration { get; set; } // Product duration
        public double m_fFxRate { get; set; } // EUR/USD
        public string m_sCurrency { get; set; } // bond currency

        public DateTime? m_FirstCouponDate { get; set; } /// first coupon date
        public double m_fCouponPerc { get; set; }
        public string m_sCouponType;
        public CashflowSchedule[] m_CashFlowSchedule; /// Cashflow schedule for +1 EUR nominal 
        public CashflowSchedule[] m_CashFlowSchedule_CPILevelBase; // CPI levels for cash flows in the Base scenario 
        public double[][] m_CashFlow_Reported; /// Local Currency Cashflow schedule for reporting purpuses 
        public double[] m_CashFlow_FXRate_Reported; /// The FX rate used to produce EUR Cashflow schedule for reporting purpuses 
        public double[][] m_CashFlow_Orig_Amonut_ToPrint; /// Local Currency Cashflow schedule for reporting purpuses 
        public DateTime[][] m_CashFlow_Orig_Date_ToPrint; /// Cashflow date shedule for reporting purpuses 
        public double[][] m_CashFlow_FXRate_Orig_ToPrint; /// The FX rate used to produce EUR Cashflow schedule for reporting purpuses 
        public int m_dCouponFrequency { get; set; }  /// number of  coupon payments per year
        public double m_fImpliedSpread { get; set; }
        public bool m_bImpliedSpreadFound = false;
        public const int m_dMaxIterations = 1000;

        // Constructors
        /// <summary>
        /// Creates new swaption instrument with supplied settings
        /// </summary>
        /// <param name="paymentType">Swaption type, ie Payer or Receiver</param>
        /// <param name="strike">Strike for swaption</param>
        /// <param name="expiry">Expiry date for option in swaption</param>
        /// <param name="maturity">Maturity date for swap in swaption</param>
        public Instrument_Bond(
            DateTime contractMaturity,
            Instrument_Bond_OriginalData originalData_Bond,
            Import_ValuationModule.CFixedIncomeCashFlowData[] FixedIncomeObj
            )
        {
            m_OriginalBond = originalData_Bond;
            m_bLookthroughBond = originalData_Bond.m_bLookthroughData;
            m_OriginalCashFlowData = FixedIncomeObj;

            m_Type = originalData_Bond.m_BondType;
            m_MaturityDate = contractMaturity;
            m_fCouponPerc = originalData_Bond.m_fCoupon;
            m_fFxRate = originalData_Bond.m_fFxRate;
            m_sCurrency = originalData_Bond.m_sCurrency;

            m_fNominal = originalData_Bond.m_fNominal; 
//            double fNominal = m_OriginalCashFlowData.m_fVolume; 
            m_sCouponType = originalData_Bond.m_sCouponType;
            m_dCouponFrequency = originalData_Bond.m_dCouponFrequency;
            if (null != m_OriginalCashFlowData)
            {
                m_CashFlowSchedule = new CashflowSchedule[2];
                double scale = 0;
                if (Math.Abs(m_fNominal) > 0)
                {
                    scale = 1.0 / m_fNominal; // official
                    //                scale = 1000 / fNominal; // test
                }
                int indx = (int) CashFlowType.RiskNeutral;
                m_CashFlowSchedule[indx] = new CashflowSchedule();
                foreach (Cashflow cf in m_OriginalCashFlowData[indx].m_CashFlowSched)
                {
                    Cashflow cf_new = new Cashflow(cf.m_Date, cf.m_fAmount * scale);
                    m_CashFlowSchedule[indx].Add(cf_new);
                }
                indx = (int)CashFlowType.RiskRente;
                m_CashFlowSchedule[indx] = new CashflowSchedule();
                foreach (Cashflow cf in m_OriginalCashFlowData[indx].m_CashFlowSched)
                {
                    Cashflow cf_new = new Cashflow(cf.m_Date, cf.m_fAmount * scale);
                    m_CashFlowSchedule[indx].Add(cf_new);
                }
            }
            else
            {
                m_CashFlowSchedule = null;
                // m_OriginalCashFlowData = .....
            }
            m_FirstCouponDate = originalData_Bond.m_dtFirstCouponDate;
            m_CashFlow_Reported = new double[2][];
            m_CashFlow_FXRate_Reported = null;
            m_CashFlow_Orig_Date_ToPrint = new DateTime[2][];
            m_CashFlow_Orig_Amonut_ToPrint = new double[2][];
            m_CashFlow_FXRate_Orig_ToPrint = new double[2][];
        }

        // Methods
        /// <summary>
        /// Name of instrument
        /// </summary>
        /// <returns>Name of instrument</returns>
        public override string Name()
        {
            return m_OriginalBond.m_sSecurityType_LL;
        }
        public static CashflowSchedule createCashFlowSchedule(DateTime dtNow, DateTime dtMaturity, double couponPerc)
        {
            CashflowSchedule schedule = new CashflowSchedule();
            if (dtMaturity < dtNow)
            {
                return schedule;
            }
            Cashflow cf;
            DateTime tDate = dtMaturity;
            cf = new Cashflow(tDate, 1.0 + couponPerc);
            schedule.Add(cf);
            if (0 != couponPerc)
            {
                while (tDate > dtNow)
                {
                    tDate = tDate.AddYears(-1);
                    if (tDate >= dtNow)
                    {
                        cf = new Cashflow(tDate, couponPerc);
                        schedule.Add(cf);
                    }
                }
            }
            return schedule;
        }

        /// <summary>
        /// Initialize swaption instrument
        /// </summary>
        /// <param name="dtNow">Valuation date</param>
        /// <param name="zeroCurve">Zero curve to determine implied volatility</param>
        /// <param name="cleanPrice">Theoretical price</param>
        public void Init(DateTime dtNow, Curve zeroSwapCurve, IndexCPI CPIindex, double marketValue)
        {
            if (null == m_CashFlowSchedule)
            {
                m_CashFlowSchedule = new CashflowSchedule[2];
                if (m_bCallable)
                {
                    m_CashFlowSchedule[(int)CashFlowType.RiskRente] = Instrument_Bond.createCashFlowSchedule(dtNow, m_MaturityDate, m_fCouponPerc);
//                    m_CashFlowSchedule[(int)CashFlowType.RiskRente] = Instrument_Bond.createCashFlowSchedule(dtNow, m_CallDate, m_fCouponPerc);
                }
                else
                {
                    m_CashFlowSchedule[(int)CashFlowType.RiskRente] = Instrument_Bond.createCashFlowSchedule(dtNow, m_MaturityDate, m_fCouponPerc);
                }

                m_CashFlowSchedule[(int)CashFlowType.RiskNeutral] = m_CashFlowSchedule[(int)CashFlowType.RiskRente];
            }
            m_CashFlowSchedule_CPILevelBase = new CashflowSchedule[2];
            int indx = (int)CashFlowType.RiskNeutral;
            m_CashFlowSchedule_CPILevelBase[indx] = new CashflowSchedule();
            if (null != CPIindex && BondType.INFLATION == m_Type)
            {
                foreach (Cashflow cf in m_CashFlowSchedule[indx])
                {
                    if (cf.m_Date >= dtNow)
                    {
                        double cpi = CPIindex.getDailyInflationReference(cf.m_Date);
                        Cashflow cf_new = new Cashflow(cf.m_Date, cpi);
                        m_CashFlowSchedule_CPILevelBase[indx].Add(cf_new);
                    }
                }
            }
            indx = (int)CashFlowType.RiskRente;
            m_CashFlowSchedule_CPILevelBase[indx] = new CashflowSchedule();
            if (null != CPIindex && BondType.INFLATION == m_Type)
            {
                foreach (Cashflow cf in m_CashFlowSchedule[indx])
                {
                    if (cf.m_Date >= dtNow)
                    {
                        double cpi = CPIindex.getDailyInflationReference(cf.m_Date);
                        Cashflow cf_new = new Cashflow(cf.m_Date, cpi);
                        m_CashFlowSchedule_CPILevelBase[indx].Add(cf_new);
                    }
                }
            }
            m_fMarketPrice = marketValue; // in EUR
            double marketValueNormalized = 0;
            if (Math.Abs(m_fNominal) > 0)
            {
                marketValueNormalized = marketValue / m_fNominal / m_fFxRate; // in issued Currency
            }
            else
            {
                m_fDuration = 0;
                m_fModelPrice = 0;
                m_fModelValueAtZeroSpread = 0;
                m_fImpliedSpread = 0;
                return;
            }
            m_fImpliedSpread = getImpliedBondSpread(dtNow, zeroSwapCurve, null, marketValueNormalized);
            if (m_bDefaulted)
            {
                m_fDuration = 0;
                m_fModelPrice = m_fMarketPrice;
                m_fModelValueAtZeroSpread = m_fMarketPrice;
            }
            else
            {
                double valueOfNominal = 1;
                m_fModelPrice = getPriceBond(dtNow, zeroSwapCurve, null, m_fImpliedSpread, out valueOfNominal);
                m_fDuration = 0;
                double mv_up = getPriceBond(dtNow, zeroSwapCurve, null, m_fImpliedSpread + 0.0001, out valueOfNominal);
                double mv_down = getPriceBond(dtNow, zeroSwapCurve, null, m_fImpliedSpread - 0.0001, out valueOfNominal);
                m_fDuration = (mv_down - mv_up) / m_fModelPrice / 0.0002;
                m_fModelPrice *= m_fNominal * m_fFxRate; // in EUR
                m_fModelValueAtZeroSpread = getPrice(dtNow, zeroSwapCurve, null, 0, m_fFxRate); // in EUR
            }
        }
        /// <summary>
        /// Calculate expiry for option in years
        /// </summary>
        /// <param name="dtNow">Valuation date</param>
        /// <returns>Expiry for option, expressed in years</returns>
        public double getMaturity(DateTime dtNow)
        {
            return getTimeToDate(dtNow, m_MaturityDate);
        }
        public double getTimeToDate(DateTime dtNow, DateTime date)
        {
            return DateTimeExtensions.YearFrac(dtNow, date, Daycount.ACT_ACT);
        }
        public void calculateStandardizedCashFlows(DateTime dtNow, Curve interestRateCurve, Curve interestRateCurveEUR, 
            DateTime[] timePoints, CashFlowType typeCF)
        {
            CashflowSchedule Schedule = m_CashFlowSchedule[(int)typeCF];
            int N = 0;
            if (m_MaturityDate >= dtNow)
            {
                foreach (Cashflow cf in Schedule)
                {
                    if (cf.m_Date >= dtNow)
                    {
                        N++;
                    }
                }
            }
            // Expected Cash Flow:
            double[] values = new double[N]; // expected cash flow
            DateTime[] dates = new DateTime[N];
            int i = 0;
            foreach (Cashflow cf in Schedule)
            {
                if (cf.m_Date >= dtNow)
                {
                    values[i] = cf.m_fAmount;
                    dates[i] = cf.m_Date;
                    i++;
                }
            }
            int K = timePoints.Length;
            // FX Rates:
            double Disc, Disc_EUR;
            m_CashFlow_FXRate_Reported = new double[K];
            double[] timePointsInYears = new double[K];
            for (int k = 0; k < K; k++)
            {
                timePointsInYears[k] = getTimeToDate(dtNow, timePoints[k]);
                Disc = interestRateCurve.DiscountFactor(timePointsInYears[k]);
                Disc_EUR = interestRateCurveEUR.DiscountFactor(timePointsInYears[k]);
                m_CashFlow_FXRate_Reported[k] = m_fFxRate * Disc / Disc_EUR;
            }
            // Cash Flow:
            double[] cashFlow = new double[K];
            i = 0;
            cashFlow[0] = 0;
            while (i < N && dates[i].CompareTo(timePoints[0]) <= 0)
            {
                cashFlow[0] += values[i];
                i++;
            }
            double alpha = 0;
            double t = 0;
            double dt = timePointsInYears[0];
            for (int k = 1; k < K; k++)
            {
                dt = timePointsInYears[k] - timePointsInYears[k-1];
                cashFlow[k] = 0;
                while (i < N && dates[i].CompareTo(timePoints[k]) <= 0)
                {
                    t = getTimeToDate(dtNow, dates[i]);
                    alpha = (timePointsInYears[k] - t) / dt;
                    cashFlow[k-1] += alpha*values[i];
                    cashFlow[k] += (1-alpha)*values[i];
                    i++;
                }
            }
            for (; i < N; i++)
            {
                cashFlow[K - 1] += values[i];
            }
            m_CashFlow_Reported[(int)typeCF] = cashFlow;
        }
        public void calculateCashFlows_Orig_ToPrint(DateTime dtNow, Curve interestRateCurve, Curve interestRateCurveEUR, CashFlowType typeCF)
        {
            CashflowSchedule Schedule = m_CashFlowSchedule[(int)typeCF];
            int N = 0;
            if (m_MaturityDate >= dtNow)
            {
                foreach (Cashflow cf in Schedule)
                {
                    if (cf.m_Date >= dtNow)
                    {
                        N++;
                    }
                }
            }
            // Expected Cash Flow:
            double[] cashFlow = new double[N]; // expected cash flow
            DateTime[] dates = new DateTime[N];
            int i = 0;
            foreach (Cashflow cf in Schedule)
            {
                if (cf.m_Date >= dtNow)
                {
                    cashFlow[i] = cf.m_fAmount;
                    dates[i] = cf.m_Date;
                    i++;
                }
            }
            // FX Rates:
            double Disc, Disc_EUR;
            double[] fxRates = new double[N];
            double[] timePointsInYears = new double[N];
            for (int k = 0; k < N; k++)
            {
                timePointsInYears[k] = getTimeToDate(dtNow, dates[k]);
                Disc = interestRateCurve.DiscountFactor(timePointsInYears[k]);
                Disc_EUR = interestRateCurveEUR.DiscountFactor(timePointsInYears[k]);
                fxRates[k] = m_fFxRate * Disc / Disc_EUR;
            }
            m_CashFlow_Orig_Date_ToPrint[(int)typeCF] = dates;
            m_CashFlow_Orig_Amonut_ToPrint[(int)typeCF] = cashFlow;
            m_CashFlow_FXRate_Orig_ToPrint[(int)typeCF] = fxRates;
        }
        /**
         * Bond Price in Issued Currency for nominal of 1 CURR unit
         */
        public double getPriceBond(DateTime dtNow, Curve zeroSwapCurve, IndexCPI CPIindex, double spread, out double valueOfNominal)
        {// in Local Currency
            Curve spreadCurve = new Curve(spread);
            Curve discountingCurve = zeroSwapCurve + spreadCurve;
            double yearFrac, df, factor, price = 0;
            int RISK = (int)CashFlowType.RiskRente;
            foreach (Cashflow cf in m_CashFlowSchedule[RISK])
            {
                if (cf.m_Date >= dtNow)
                {
                    yearFrac = DateTimeExtensions.YearFrac(dtNow, cf.m_Date, Daycount.ACT_365);
                    df = discountingCurve.DiscountFactor(yearFrac);
                    factor = 1;
                    if (BondType.INFLATION == m_Type)
                    {
                        if (null != m_CashFlowSchedule_CPILevelBase && null != CPIindex)
                        {
                            double cpi_base = m_CashFlowSchedule_CPILevelBase[RISK].m_cashflows[cf.m_Date].m_fAmount;
                            double cpi = CPIindex.getDailyInflationReference(cf.m_Date);
                            factor = cpi / cpi_base;
                        }
                        else
                        {
                            factor = 1; // not provided index
                        }
                    }
                    
                    price += cf.m_fAmount * df * factor;
                }
            }
            yearFrac = DateTimeExtensions.YearFrac(dtNow, m_MaturityDate, Daycount.ACT_365);
            valueOfNominal = discountingCurve.DiscountFactor(yearFrac);
            return price;
        }
        public double getImpliedBondSpread(DateTime dtNow, Curve zeroCurve, IndexCPI CPIindex, double price)
        {
            m_bImpliedSpreadFound = false;
            m_bDefaulted = false;
            double MaxError = 0.01; // 1 cent
            double maxSpread = 1.0; // +100%
            double minSpread = -0.1; // -10%
            double lowerSpread = -0.1;
            double upperSpread = 0.1;
            double spread = 0.00;
            double valueOfNominal = 1;
            double p = getPriceBond(dtNow, zeroCurve, CPIindex, spread, out valueOfNominal);
            int z = m_fCouponPerc > 0 ? 1 : -1;
            double valueOfCoupons = p - valueOfNominal;
            double valueOfCouponsAbs = z * valueOfCoupons;
            double priceAbs = Math.Abs(price);
            int zp = price > 0 ? 1 : -1;
            if (p < price)
            {
                while (p < price)
                {
                    if (spread <= minSpread)
                    {
                        return minSpread;
                    }
                    upperSpread = spread;
                    spread -= 0.002;
                    p = getPriceBond(dtNow, zeroCurve, CPIindex, spread, out valueOfNominal);
                }
                lowerSpread = spread;
            }
            else if (p > price)
            {
                while (p > price)
                {
                    if (spread >= maxSpread)
                    {
                        m_bDefaulted = true;
                        return maxSpread;
                    }
                    lowerSpread = spread;
                    spread += 0.002;
                    p = getPriceBond(dtNow, zeroCurve, CPIindex, spread, out valueOfNominal);
                }
                upperSpread = spread;
            }
            else
            {
                m_bImpliedSpreadFound = true;
                return spread;
            }

            spread = lowerSpread + (upperSpread - lowerSpread) / 2;
            m_dIterations = 0;
            p = getPriceBond(dtNow, zeroCurve, CPIindex, spread, out valueOfNominal);
            MaxError = MaxError / Math.Abs(m_fNominal);
            while ((Math.Abs(p - price) > MaxError) && (m_dIterations < m_dMaxIterations))
            {
                m_dIterations++;
                if (p > price)
                {
                    lowerSpread = spread;
                }
                else if (p < price)
                {
                    upperSpread = spread;
                }
                spread = lowerSpread + (upperSpread - lowerSpread) / 2;
                p = getPriceBond(dtNow, zeroCurve, CPIindex, spread, out valueOfNominal);
            }
            if (m_dIterations < m_dMaxIterations)
            {
                m_bImpliedSpreadFound = true;
            }
            return spread;
        }
        /**
         * Bond price in EUR
         */
        public double getPrice(DateTime dtNow, Curve zeroSwapCurve, IndexCPI CPIindex,  double spreadBond, double fxRate)
        {// in EUR
            if (m_bDefaulted || !m_bImpliedSpreadFound || 0 == m_fNominal)
            {
                return m_fModelPrice * fxRate / m_fFxRate;
            }
            double valueOfNominal = 1;
            double priceBond = getPriceBond(dtNow, zeroSwapCurve, CPIindex, spreadBond, out valueOfNominal);
            return priceBond * m_fNominal * fxRate; // in EUR
        }
        public double getPrice(DateTime dtNow, Curve zeroSwapCurve)
        {
            if (m_bDefaulted || !m_bImpliedSpreadFound || 0 == m_fNominal)
            {
                return m_fModelPrice;
            }
            double priceBond = getPrice(dtNow, zeroSwapCurve, null, m_fImpliedSpread, m_fFxRate);
            return priceBond;
        }
    }

}
