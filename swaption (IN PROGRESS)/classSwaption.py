import functionsSwaption as fs

class Swaption:

    def __init__(self, cicCode, securityId, marketValue, nominalValue, volatilityPercentageIMW, strikePercentage, dateSwapExpiry, dateSwaptionExpiry):
        self.cicCode = cicCode
        self.securityId = securityId
        self.marketValue = marketValue
        self.nominalValue = nominalValue
        self.volatilityPercentageIMW = volatilityPercentageIMW
        self.strikePercentage = strikePercentage
        self.dateSwapExpiry = dateSwapExpiry
        self.dateSwaptionExpiry = dateSwaptionExpiry
        self.cashSettled = True # (voorlopige) aanname in TRT is dat alle swaptions cash settled zijn
        self.getSwaptionType()

    def getSwaptionType(self):
        if self.cicCode[2:4] == 'B6':
            self.swaptionType = 'call'
        elif self.cicCode[2:4] == 'C6':
            self.swaptionType = 'put'
        else:
            raise ValueError('Swaption type kon niet worden gevonden')
    
    def getImpliedVolatility(self, curve):
        self.impliedVolatility = 2 * fs.getPrice(self.marketValue)

    def getPrice(self, curve):
        price = fs.getPrice(self.nominalValue)
        return price
