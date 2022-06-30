using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TotalRisk.ExcelWrapper;
using TotalRisk.Utilities;

namespace TotalRisk.ValuationModule
{

    public enum SwaptionType
    {
        Payer,
        Receiver
    }

    public class Instrument_Swaption_OriginalData
    {
        public string m_sSelectieIndex_LL;
        public string m_sScope3;
        public string m_sPortfolioID;
        public string m_sInstrumentType;
        public string m_sCICLL;
        public string m_sSecurityID;
        public string m_sSecurityName;
        public string m_sType;
        public string m_sCurrency = "EUR";
        public double m_fSwaptionExpiry;
        public double m_fSwapExpiry;
        public double m_fStrike;
        public double m_fSwaptionVolatility;
        public double m_fFxRate;
        public double m_fNominal;
        public double m_fMarketValue;
        public string m_sGroupCounterpartyName;
        public string m_sGroupCounterpartyLEI;
        public string m_sGroupCounterpartyCQS;
        public int m_dGroupCounterpartyCQS;
    }

    /// <summary>
    /// This class define a swaption instrument
    /// </summary>
    public class Instrument_Swaption : FinancialInstrument
    {
        /****  MODEL VERSION NUMBER   ******/
        static public string getVersion()
        {

            //            string version = "Version 5.00: created in 2017"; //    
//            string version = "Version 5.10: created in 2019-02-06"; //    
            string version = "Version 5.08: created in 2019-02-02 : Bachelier-1900 normal model: SCR floor is removed"; //    
            return version;
        }
        /*******  DATA  ********/
        public Instrument_Swaption_OriginalData m_OriginalData;

        public const double CONSTL = 0.3989422804; // 1/Sqrt(2 Pi)
        // Properties
        public bool m_bCashSettled { get; set; } // whether it is cash or physical settled
        public SwaptionType m_Type { get; set; } /// Swaption type, ie Payer or Receiver
        public double m_fStrike { get; set; } /// Strike for swaption
        public DateTime m_ExpiryDate { get; set; } /// Expiry date for option in swaption
        public DateTime m_MaturityDate { get; set; } /// Expiry date for swap in swaption
        public double m_fMarketPrice { get; set; } /// Market price
        public double m_fModelPrice { get; set; } /// Theoretical price
        public double m_fCleanPriceOnePercentVol { get; set; } /// Theoretical price with one percent vol
        public double m_fIntrisicValue { get; set; } /// Theoretical price with zero percent vol
        public double m_fImpliedVol { get; set; } /// Implied volatility
        public double m_fVolatility { get; set; } /// actual volatility
        public double m_fForwardRate { get; set; } /// forward rate
        public string m_sCommentOnMoneyness;
        public const int m_dMaxIterations = 1000;
        public Boolean m_bNormalVol = true;

        // Constructors
        /// <summary>
        /// Creates new swaption instrument with supplied settings
        /// </summary>
        /// <param name="type">Swaption type, ie Payer or Receiver</param>
        /// <param name="strike">Strike for swaption</param>
        /// <param name="expiry">Expiry date for option in swaption</param>
        /// <param name="maturity">Maturity date for swap in swaption</param>
        public Instrument_Swaption(Instrument_Swaption_OriginalData OriginalData, SwaptionType type, bool cashSettled, double strike, DateTime expiry, DateTime maturity)
        {
            m_OriginalData = OriginalData;
            this.m_Type = type;
            this.m_fStrike = strike;
            this.m_ExpiryDate = expiry;
            this.m_MaturityDate = maturity;
            this.m_bCashSettled = cashSettled;
        }

        // Methods
        /// <summary>
        /// Name of instrument
        /// </summary>
        /// <returns>Name of instrument</returns>
        public override string Name()
        {
            return "Swaption";
        }

        /// <summary>
        /// Initialize swaption instrument
        /// </summary>
        /// <param name="dtNow">Valuation date</param>
        /// <param name="zeroCurve">Zero curve to determine implied volatility</param>
        /// <param name="cleanPrice">Theoretical price</param>
        public void Init_Ad_Hoc(DateTime dtNow, Curve zeroCurve, double marketPrice, double volatility)
        {
            m_fMarketPrice = marketPrice;
            m_fVolatility = volatility;

            m_fForwardRate = getForwardRate(dtNow, zeroCurve);
            m_fModelPrice = getPrice(dtNow, zeroCurve, volatility);
            if (m_bNormalVol)
            {
                m_fImpliedVol = getImpliedVolatility(dtNow, zeroCurve, marketPrice);
            }
            else if (m_fStrike > 0.01)
            {
                m_fImpliedVol = getImpliedVolatility(dtNow, zeroCurve, marketPrice);
            }
            else
            {
                m_fImpliedVol = 0;
            }
            m_fIntrisicValue = getPrice(dtNow, zeroCurve, 0);

            m_fCleanPriceOnePercentVol = getPrice(dtNow, zeroCurve, 0.01);
            if (m_fForwardRate <= 0)
            {
                if (SwaptionType.Payer == m_Type)
                {
                    m_sCommentOnMoneyness = "OTM";
                }
                else
                {
                    m_sCommentOnMoneyness = "ITM";
                }
            }
            if (m_fStrike > 1.05 * m_fForwardRate)
            {
                if (SwaptionType.Payer == m_Type)
                {
                    m_sCommentOnMoneyness = "OTM";
                }
                else
                {
                    m_sCommentOnMoneyness = "ITM";
                }
            }
            else if (m_fStrike < 0.95 * m_fForwardRate)
            {
                if (SwaptionType.Payer == m_Type)
                {
                    m_sCommentOnMoneyness = "ITM";
                }
                else
                {
                    m_sCommentOnMoneyness = "OTM";
                }
            }
            else
            {
                m_sCommentOnMoneyness = "ATM";
            }

        }
        public void Init(DateTime dtNow, Curve zeroCurve, double marketPrice, double volatility)
        {
            m_fMarketPrice = marketPrice;
            m_fVolatility = volatility;

            m_fForwardRate = getForwardRate(dtNow, zeroCurve);
            m_fModelPrice = getPrice(dtNow, zeroCurve, volatility);
            if (m_bNormalVol)
            {
                m_fImpliedVol = getImpliedVolatility(dtNow, zeroCurve, marketPrice);
            }
            else if (m_fStrike > 0.01)
            {
                m_fImpliedVol = getImpliedVolatility(dtNow, zeroCurve, marketPrice);
            }
            else
            {
                m_fImpliedVol = 0;
            }
            m_fIntrisicValue = getPrice(dtNow, zeroCurve, 0);

            m_fCleanPriceOnePercentVol = getPrice(dtNow, zeroCurve, 0.01);
            if (m_fForwardRate <= 0)
            {
                if (SwaptionType.Payer == m_Type)
                {
                    m_sCommentOnMoneyness = "OTM";
                }
                else
                {
                    m_sCommentOnMoneyness = "ITM";
                }
            }
            if (m_fStrike > 1.05 * m_fForwardRate)
            {
                if (SwaptionType.Payer == m_Type)
                {
                    m_sCommentOnMoneyness = "OTM";
                }
                else
                {
                    m_sCommentOnMoneyness = "ITM";
                }
            }
            else if (m_fStrike < 0.95 * m_fForwardRate)
            {
                if (SwaptionType.Payer == m_Type)
                {
                    m_sCommentOnMoneyness = "ITM";
                }
                else
                {
                    m_sCommentOnMoneyness = "OTM";
                }
            }
            else
            {
                m_sCommentOnMoneyness = "ATM";
            }

        }
        public void Init(DateTime dtNow, Curve zeroCurve, double marketPrice, double hullWhiteA, double hullWhiteSigma)
        {
            m_fMarketPrice = marketPrice;
            m_fVolatility = hullWhiteSigma;

            m_fForwardRate = getForwardRate(dtNow, zeroCurve);
            m_fModelPrice = getPrice(dtNow, zeroCurve, hullWhiteA, hullWhiteSigma);
            m_fImpliedVol = 0;
            m_fIntrisicValue = getPrice(dtNow, zeroCurve, 0);

            m_fCleanPriceOnePercentVol = getPrice(dtNow, zeroCurve, hullWhiteA, 0.01);
            if (m_fForwardRate <= 0)
            {
                if (SwaptionType.Payer == m_Type)
                {
                    m_sCommentOnMoneyness = "OTM";
                }
                else
                {
                    m_sCommentOnMoneyness = "ITM";
                }
            }
            if (m_fStrike > 1.05 * m_fForwardRate)
            {
                if (SwaptionType.Payer == m_Type)
                {
                    m_sCommentOnMoneyness = "OTM";
                }
                else
                {
                    m_sCommentOnMoneyness = "ITM";
                }
            }
            else if (m_fStrike < 0.95 * m_fForwardRate)
            {
                if (SwaptionType.Payer == m_Type)
                {
                    m_sCommentOnMoneyness = "ITM";
                }
                else
                {
                    m_sCommentOnMoneyness = "OTM";
                }
            }
            else
            {
                m_sCommentOnMoneyness = "ATM";
            }

        }

        /// <summary>
        /// Calculate tenor for swap in years
        /// </summary>
        /// <returns>Tenor for swap, expressed in years</returns>
        public int getTenor()
        {
            return m_MaturityDate.Year - m_ExpiryDate.Year;
        }

        /// <summary>
        /// Calculate expiry for option in years
        /// </summary>
        /// <param name="dtNow">Valuation date</param>
        /// <returns>Expiry for option, expressed in years</returns>
        public double getMaturity(DateTime dtNow)
        {
            return DateTimeExtensions.YearFrac(dtNow, m_ExpiryDate, Daycount.ACT_ACT);
        }

        /// <summary>
        /// Calculate forward rate for swaption
        /// </summary>
        /// <param name="dtNow">Valuation date</param>
        /// <param name="zeroCurve">Zero curve</param>
        /// <param name="vol">Volatility</param>
        /// <returns>Forward rate for swaption</returns>
        public double getForwardRate(DateTime dtNow, Curve zeroCurve)
        {
            int tenor = getTenor();
            double expiry = getMaturity(dtNow);
            double sumDf = 0;
            double[] dfSwap = new double[2 * tenor + 1];
            for (int idx = 0; idx < dfSwap.Length; idx++)
            {
                dfSwap[idx] = zeroCurve.DiscountFactor(expiry + 0.5 * idx);
                if (idx > 0)
                {
                    sumDf += dfSwap[idx] * 0.5;
                }
            }
            double fwd_rate = (dfSwap[0] - dfSwap[dfSwap.Length - 1]) / sumDf;
            return fwd_rate;
        }

        /// <summary>
        ///  This function calculates the price of a swaption with notional of 1 EUR
        /// </summary>
        /// <param name="dtNow">Valuation date</param>
        /// <param name="zeroSwapCurve">Zero curve</param>
        /// <param name="volatility">the volatility of the swaption</param>
        /// <returns>Price for swaption</returns>

        public double getPrice(DateTime dtNow, Curve zeroSwapCurve, Curve zeroEONIA_Curve, double volatility)
        {
            double eps = 0.00001; // 0.1 bp
            double strike = m_fStrike;
            if (m_bCashSettled || !m_bNormalVol)
            {
                if (Math.Abs(m_fStrike) < eps)
                {
                    strike = m_fStrike < 0 ? -eps : eps;
                }
            }
            int z = 1;	// payer
            if (SwaptionType.Receiver == m_Type)
            {
                z = -1; // receiver
            }
            int tenor = getTenor();
            double expiry = getMaturity(dtNow);

            double sumDf = 0;
            double sumDf_EONIA = 0;
            double sumDf_EONIA_c = 0;
            double[] dfSwap = new double[2 * tenor + 1];
            double[] dfSwap_EONIA = new double[2 * tenor + 1];
            for (int idx = 0; idx < dfSwap.Length; idx++)
            {
                dfSwap[idx] = zeroSwapCurve.DiscountFactor(expiry + 0.5 * idx);
                dfSwap_EONIA[idx] = zeroEONIA_Curve.DiscountFactor(expiry + 0.5 * idx);
                if (idx > 0)
                {
                    double c = dfSwap[idx-1] / dfSwap[idx] - 1;
                    sumDf += dfSwap[idx] * 0.5;
                    sumDf_EONIA_c += dfSwap_EONIA[idx] * c;
                    sumDf_EONIA += dfSwap_EONIA[idx] * 0.5;
                }
            }
            double fwd_rate = sumDf_EONIA_c / sumDf_EONIA;
            if (m_bCashSettled || !m_bNormalVol)
            {
                if (Math.Abs(fwd_rate) < eps)
                {
                    fwd_rate = fwd_rate < 0 ? -eps : eps;
                }
            }
            double price;
            if (volatility <= 0)
            {
                price = z * (fwd_rate - strike);
            }
            else
            {
                if (m_bNormalVol)
                {
                    double d = (fwd_rate - strike) / volatility / Math.Sqrt(expiry);
                    double Nd = Statistics.CND(z * d);
                    price = z * (fwd_rate - strike) * Nd + volatility * Math.Sqrt(expiry) * CONSTL * Math.Exp(-d * d / 2);
                }
                else
                {
                    double d1 = (Math.Log(fwd_rate / strike) + volatility * volatility / 2 * expiry) / volatility / Math.Sqrt(expiry);
                    double d2 = d1 - volatility * Math.Sqrt(expiry);
                    double Nd_1 = Statistics.CND(z * d1);
                    double Nd_2 = Statistics.CND(z * d2);
                    price = z * (fwd_rate * Nd_1 - strike * Nd_2);
                }
            }
            if (price < 0)
            {
                price = 0;
            }
            if (m_bCashSettled)
            { // Haug formula for cash sattled swaptions assumes that 6 month compounded swap rate used as the discounting rate.
                double factor = (1 - Math.Pow(1 + fwd_rate / 2, -2 * tenor)) / fwd_rate;
                factor *= dfSwap[0];
                return price * factor;
            }
            else
            {
                return price * sumDf_EONIA;
            }
        }
        public double getPrice(DateTime dtNow, Curve zeroCurve, double volatility)
        {
            double eps = 0.00001; // 0.1 bp
            double strike = m_fStrike;
            if (m_bCashSettled || !m_bNormalVol)
            {
                if (Math.Abs(m_fStrike) < eps)
                {
                    strike = m_fStrike < 0 ? -eps : eps;
                }
            }
            int z = 1;	// payer
            if (SwaptionType.Receiver == m_Type)
            {
                z = -1; // receiver
            }
            int tenor = getTenor();
            double expiry = getMaturity(dtNow);

            double sumDf = 0;
            double[] dfSwap = new double[2 * tenor + 1];
            for (int idx = 0; idx < dfSwap.Length; idx++)
            {
                dfSwap[idx] = zeroCurve.DiscountFactor(expiry + 0.5 * idx);
                if (idx > 0)
                {
                    sumDf += dfSwap[idx] * 0.5;
                }
            }
            double fwd_rate = (dfSwap[0] - dfSwap[dfSwap.Length - 1]) / sumDf;
            if (m_bCashSettled || !m_bNormalVol)
            {
                if (Math.Abs(fwd_rate) < eps)
                {
                    fwd_rate = fwd_rate < 0 ? -eps : eps;
                }
            }
            double price;
            if (volatility <= 0)
            {
                price = z * (fwd_rate - strike);
            }
            else
            {
                if (m_bNormalVol)
                {
                    double d = (fwd_rate - strike) / volatility / Math.Sqrt(expiry);
                    double Nd = Statistics.CND(z * d);
                    price = z * (fwd_rate - strike) * Nd + volatility * Math.Sqrt(expiry) * CONSTL * Math.Exp(-d * d / 2);
                }
                else
                {
                    double d1 = (Math.Log(fwd_rate / strike) + volatility * volatility / 2 * expiry) / volatility / Math.Sqrt(expiry);
                    double d2 = d1 - volatility * Math.Sqrt(expiry);
                    double Nd_1 = Statistics.CND(z * d1);
                    double Nd_2 = Statistics.CND(z * d2);
                    price = z * (fwd_rate * Nd_1 - strike * Nd_2);
                }
            }
            if (price < 0)
            {
                price = 0;
            }
            if (m_bCashSettled)
            { // Haug formula for cash sattled swaptions assumes that 6 month compounded swap rate used as the discounting rate.
                double factor = (1 - Math.Pow(1 + fwd_rate / 2, -2 * tenor)) / fwd_rate;
                factor *= dfSwap[0];
                return price * factor;
            }
            else
            {
                return price * sumDf;
            }
        }
        public double getPrice_OLD(DateTime dtNow, Curve zeroCurve, double volatility)
        {
            double strike = m_fStrike > 0.0001 ? m_fStrike : 0.0001; // to manage zero strile swaptions: minimum is 0.1%
            int z = 1;	// payer
            if (SwaptionType.Receiver == m_Type)
            {
                z = -1; // receiver
            }
            int tenor = getTenor();
            double expiry = getMaturity(dtNow);

            double sumDf = 0;
            double[] dfSwap = new double[2 * tenor + 1];
            for (int idx = 0; idx < dfSwap.Length; idx++)
            {
                dfSwap[idx] = zeroCurve.DiscountFactor(expiry + 0.5 * idx);
                if (idx > 0)
                {
                    sumDf += dfSwap[idx] * 0.5;
                }
            }
            double fwd_rate = (dfSwap[0] - dfSwap[dfSwap.Length - 1]) / sumDf;
            if (fwd_rate < 0.00001) fwd_rate = 0.00001; // Min 0.1 bp
            if (fwd_rate <= 0)
            {
                if (SwaptionType.Payer == m_Type)
                {
                    return 0;
                }
                else
                {
                    return strike * sumDf; // just swap price
                }
            }
            double price = 0;
            if (volatility <= 0)
            {
                price = z * (fwd_rate - strike);
            }
            else
            {
                if (m_bNormalVol)
                {
                    double d = (fwd_rate - strike) / volatility / Math.Sqrt(expiry);
                    double Nd = Statistics.CND(z * d);
                    price = z * (fwd_rate - strike) * Nd + volatility * Math.Sqrt(expiry) * CONSTL * Math.Exp(-d * d / 2);
                }
                else
                {
                    double d1 = (Math.Log(fwd_rate / strike) + volatility * volatility / 2 * expiry) / volatility / Math.Sqrt(expiry);
                    double d2 = d1 - volatility * Math.Sqrt(expiry);
                    double Nd_1 = Statistics.CND(z * d1);
                    double Nd_2 = Statistics.CND(z * d2);
                    price = z * (fwd_rate * Nd_1 - strike * Nd_2);
                }
            }
            if (price < 0)
            {
                price = 0;
            }
            if (m_bCashSettled)
            { // Haug formula for cash sattled swaptions assumes that 6 month compounded swap rate used as the discounting rate.
                double factor = (1 - Math.Pow(1 + fwd_rate / 2, -2 * tenor)) / fwd_rate;
                factor *= dfSwap[0];
                return price * factor;
            }
            else
            {
                return price * sumDf;
            }
        }
        /// <summary>
        /// Calculate volatility for swaption
        /// </summary>
        /// <param name="dtNow">Valuation date</param>
        /// <param name="zeroCurve">Zero curve used to determine price</param>
        /// <param name="price">Theoretial price</param>
        /// <returns>Volatility for swaption</returns>
        /// <remarks>This function uses a goalseek implementation to determine the implied volatility</remarks>
        public double getImpliedVolatility(DateTime dtNow, Curve zeroCurve, double price)
        {
            const double MaxError = 1e-12;
            double forwardRate = getForwardRate(dtNow, zeroCurve);
            if (forwardRate <= 0)
            {
                if (SwaptionType.Payer == m_Type)
                {
                    return m_fVolatility;
                }
                else
                {
                    return m_fVolatility;
                }
            }
            double lowerVol = 0;
            double upperVol = 1;
            double p = getPrice(dtNow, zeroCurve, upperVol);
            int iter = 0;
            while (p < price)
            {
                iter++;
                upperVol *= 2;
                p = getPrice(dtNow, zeroCurve, upperVol);
                if (iter > 100)
                {
                    iter += 0;
                }
            }
            double vol = lowerVol + (upperVol - lowerVol) / 2;
            m_dIterations = 0;
            p = getPrice(dtNow, zeroCurve, vol);
            while ((Math.Abs(p - price) > MaxError) && (m_dIterations < m_dMaxIterations))
            {
                m_dIterations++;
                if (p > price)
                {
                    upperVol = vol;
                }
                else if (p < price)
                {
                    lowerVol = vol;
                }
                vol = lowerVol + (upperVol - lowerVol) / 2;
                p = getPrice(dtNow, zeroCurve, vol);
            }
            if (vol < MaxError)
            {
                vol = 0;
            }
            return vol;
        }
        public double getPrice(DateTime dtNow, Curve zeroCurve, double hullWhiteA, double hullWhiteSigma)
        {
            int z = 1; // recever
            if (SwaptionType.Payer == m_Type)
            {
                z = -1;
            }
            double X = m_fStrike;
            double T = getMaturity(dtNow);
            int frequency = 2;
            int tenorYears = getTenor();
            int N = frequency * tenorYears;
            double a = hullWhiteA;
            double sigma = hullWhiteSigma;
            double[] Si = new double[N];
            double[] DS = new double[N];
            double[] c = new double[N];
            double DT = zeroCurve.DiscountFactor(T);
            for (int i = 0; i < N; i++)
            {
                c[i] = X / frequency;
                Si[i] = T + (i + 1.0) / frequency;
                DS[i] = zeroCurve.DiscountFactor(Si[i]);
            }
            c[N - 1] += 1.0;
            double shift = 0.001;
            double market_instantaneous_forward_rate = (
                        zeroCurve.GetRate(T + shift) * (T + shift)
                      - zeroCurve.GetRate(T - shift) * (T - shift)
                    ) / (2 * shift);

            double R = calculateSpotRate(a, sigma, market_instantaneous_forward_rate, T, Si, DT, DS, c);
            double price = getPriceAnalytical(z, a, sigma, T, Si, market_instantaneous_forward_rate, DT, DS, c, R);
            return price;
        }

        public static double getPriceAnalytical(
                int z,
                double a,
                double sigma,
                double T,
                double[] S,
                double fT,
                double DT,
                double[] DS,
                double[] c,
                double R
                )
        {
            int n = S.Length;
            double A;
            double B;
            double Xi;
            double SigmaP;
            double H;
            double ZBO;
            double price = 0;
            double e = Math.Exp(-2 * a * T);
            for (int i = 0; i < n; i++)
            {
                B = (1 - Math.Exp(-a * (S[i] - T))) / a;
                A = DS[i] / DT * Math.Exp(B * (fT - sigma * sigma * (1 - e) * B / 4 / a));
                Xi = A * Math.Pow(R, B);
                SigmaP = getSigmaP(a, sigma, T, S[i]);
                H = getH(Xi, SigmaP, DT, DS[i]);
                ZBO = z * (DS[i] * Statistics.CND(z * H) - DT * Xi * Statistics.CND(z * (H - SigmaP)));
                price += c[i] * ZBO;
            }
            return price;
        }

        public static double getH(
                double X,
                double SigmaP,
                double DT,
                double DS
                )
        {
            double H = Math.Log(DS / DT / X) / SigmaP + SigmaP / 2;
            return H;
        }
        public static double getSigmaP(
                double a,
                double sigma,
                double T,
                double S
                )
        {
            double e = Math.Exp(-2 * a * T);
            double B = (1 - Math.Exp(-a * (S - T))) / a;
            return B * sigma * Math.Sqrt((1 - e) / 2 / a);
        }

        public static double calculateSpotRate(
                double a,
                double sigma,
                double market_instantaneous_forward_rate,
                double T,
                double[] S,
                double DT,
                double[] DS,
                double[] c
                )
        {
            double R = 0.9;
            double R_min = 0.0;
            double R_max = 1.0;
            double value_min = 0;
            double value_max = 0;
            double value = getG(a, sigma, T, S, market_instantaneous_forward_rate, DT, DS, c, R);
            if (value > 0)
            {
                R_max = R;
                value_max = value;
                R_min = R_max;
                value_min = value_max;
                while (value_min > 0)
                {
                    R_max = R_min;
                    value_max = value_min;
                    R_min /= 2;
                    value_min = getG(a, sigma, T, S, market_instantaneous_forward_rate, DT, DS, c, R_min);
                }
            }
            else if (value < 0)
            {
                R_min = R;
                value_min = value;
                R_max = R_min;
                value_max = value_min;
                while (value_max < 0)
                {
                    R_min = R_max;
                    value_min = value_max;
                    R_max *= 2;
                    value_max = getG(a, sigma, T, S, market_instantaneous_forward_rate, DT, DS, c, R_max);
                }
            }
            else
            {
                return R;
            }

            //        R = R_min - value_min*(R_max - R_min)/(value_max - value_min);
            R = R_min + 0.5 * (R_max - R_min);
            value = getG(a, sigma, T, S, market_instantaneous_forward_rate, DT, DS, c, R);
            int nSteps = 0;
            double error = 1.0e-10;
            int maximumNumberOfSteps = 1000;
            while (Math.Abs(value) > error && nSteps < maximumNumberOfSteps)
            {
                if (value > 0)
                {
                    R_max = R;
                    value_max = value;
                }
                else
                {
                    R_min = R;
                    value_min = value;
                }
                //            R = R_min - value_min*(R_max - R_min)/(value_max - value_min);
                R = R_min + 0.5 * (R_max - R_min);
                value = getG(a, sigma, T, S, market_instantaneous_forward_rate, DT, DS, c, R);
                nSteps++;
            }
            return R;
        }
        public static double getG(
                double a,
                double sigma,
                double T,
                double[] S,
                double fT,
                double DT,
                double[] DS,
                double[] c,
                double R // ln(-R)
                )
        {
            int n = S.Length;
            double G = 0;
            double A;
            double B;
            double Xi;
            double e = Math.Exp(-2 * a * T);
            for (int i = 0; i < n; i++)
            {
                B = (1 - Math.Exp(-a * (S[i] - T))) / a;
                A = DS[i] / DT * Math.Exp(B * (fT - sigma * sigma * (1 - e) * B / 4 / a));
                Xi = A * Math.Pow(R, B);
                G += c[i] * Xi;
            }
            return (G - 1.0);
        }

        /// <summary>
        /// Calculate price for swaption
        /// </summary>
        /// <param name="dtNow">Valuation date</param>
        /// <param name="zeroCurve">Zero curve</param>
        /// <returns>Price for swaption</returns>
        public double getPrice(DateTime dtNow, Curve zeroCurve)
        {
            return getPrice(dtNow, zeroCurve, m_fVolatility);
        }
    }

}
