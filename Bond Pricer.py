import sys  # used for exit
import pandas as pd  # used for read_excel
import os  # used to check file existence

########################################################################################################################

class BOND_PRICER:

    def __init__(self, bond_type, company, maturity, coupon_rate_type=None, coupon_rate_or_margin=None, coupon_frequency=None):
        self.bond_type = bond_type  # Type of Bond: bullet, zero coupon, fixed annuities, constant amortizations, equal series repayment
        self.company = company
        self.maturity = maturity  # Maturity in years
        self.coupon_rate_type = coupon_rate_type  # "Variable" or "Fixed"
        self.coupon_rate_or_margin = coupon_rate_or_margin
        self.coupon_frequency = coupon_frequency  # Number of coupons in a year

    def price(self, risk_free_curve, spread_curve, libor_curve=None):
        if self.bond_type == 'bullet':
            return self.price_bullet(risk_free_curve, spread_curve, libor_curve)
        elif self.bond_type == 'zero coupon':
            return self.price_zero_coupon(risk_free_curve, spread_curve)
        elif self.bond_type == 'fixed annuities':
            return self.price_fixed_annuities(risk_free_curve, spread_curve, libor_curve)
        elif self.bond_type == 'constant amortizations':
            return self.price_constant_amortizations(risk_free_curve, spread_curve, libor_curve)
        elif self.bond_type == 'equal series repayment':
            return self.price_equal_series(risk_free_curve, spread_curve, libor_curve)
        else:
            raise ValueError("Invalid bond type")

    def price_bullet(self, risk_free_curve, spread_curve, libor_curve):
        price = 0
        first_coupon = self.maturity - int(self.maturity) + 1 / self.coupon_frequency

        for i in (first_coupon + j * (1 / self.coupon_frequency) for j in range(int((self.maturity - first_coupon) * self.coupon_frequency) + 1)):
            if self.coupon_rate_type == "Fixed":
                price += (self.coupon_rate_or_margin / self.coupon_frequency) / ((1 + risk_free_curve.interpolate(i) + spread_curve.interpolate(i)/10000) ** i)
            elif self.coupon_rate_type == "Variable":
                price += ((self.coupon_rate_or_margin/10000 + libor_curve.interpolate(i)) / self.coupon_frequency) / ((1 + risk_free_curve.interpolate(i) + spread_curve.interpolate(i)/10000) ** i)

        return price + 1 / ((1 + risk_free_curve.interpolate(self.maturity) + spread_curve.interpolate(self.maturity)/10000) ** self.maturity)

    def price_zero_coupon(self, risk_free_curve, spread_curve):
        return 100 / ((1 + risk_free_curve.interpolate(self.maturity) + spread_curve.interpolate(self.maturity)/10000) ** self.maturity)

    def price_fixed_annuities(self, risk_free_curve, spread_curve, libor_curve):
        price = 0
        annuity = self.calculate_annuity()
        first_coupon = self.maturity - int(self.maturity) + 1 / self.coupon_frequency

        for i in (first_coupon + j * (1 / self.coupon_frequency) for j in range(int((self.maturity - first_coupon) * self.coupon_frequency) + 1)):
            discount_rate = (1 + risk_free_curve.interpolate(i) + spread_curve.interpolate(i)/10000) ** i
            price += annuity / discount_rate

        return price

    def price_constant_amortizations(self, risk_free_curve, spread_curve, libor_curve):
        price = 0
        first_coupon = self.maturity - int(self.maturity) + 1 / self.coupon_frequency
        principal_payment = 100 / (self.maturity - first_coupon + 1)
        total_payments = int(self.maturity * self.coupon_frequency)

        for i in (first_coupon + j * (1 / self.coupon_frequency) for j in range(int((self.maturity - first_coupon) * self.coupon_frequency) + 1)):
            interest_payment = 100 * (self.maturity * self.coupon_frequency - (i - 1)) / total_payments * (self.coupon_rate_or_margin / self.coupon_frequency)
            discount_rate = (1 + risk_free_curve.interpolate(i) + spread_curve.interpolate(i)/10000) ** i
            price += (principal_payment + interest_payment) / discount_rate

        return price
    
    def price_equal_series(self, risk_free_curve, spread_curve, libor_curve):
        price = 0
        amortization = 0
        first_coupon = self.maturity - int(self.maturity) + 1
        first_coupon_with_amortization = self.maturity - int(self.maturity) + 1 / self.coupon_frequency
        total_payments = int(self.maturity * self.coupon_frequency)
        series_payment = 100 / total_payments

        for i in (first_coupon + j for j in range(int(self.maturity - first_coupon) + 1)):
            outstanding_capital = 100 - amortization

            interest_payment = outstanding_capital * (self.coupon_rate_or_margin / self.coupon_frequency)
            total_payment = interest_payment
            for j in range(total_payments):
                if i == first_coupon_with_amortization + j * (1 / self.coupon_frequency):
                    total_payment = interest_payment + series_payment
                    amortization += series_payment

            discount_rate = (1 + risk_free_curve.interpolate(i) + spread_curve.interpolate(i) / 10000) ** i
            price += total_payment / discount_rate

        return price

    def calculate_annuity(self):
        rate_per_period = self.coupon_rate_or_margin / self.coupon_frequency
        periods = self.maturity * self.coupon_frequency
        annuity = 100 * rate_per_period / (1 - (1 + rate_per_period) ** -periods)
        return annuity

    def duration(self, risk_free_curve, spread_curve, libor_curve=None):
        price = self.price(risk_free_curve, spread_curve, libor_curve)
        duration = 0
        first_coupon = self.maturity - int(self.maturity) + 1 / self.coupon_frequency

        for i in (first_coupon + j * (1 / self.coupon_frequency) for j in range(int((self.maturity - first_coupon) * self.coupon_frequency) + 1)):
            cash_flow = self.calculate_cash_flow(i, risk_free_curve, spread_curve, libor_curve)
            discount_rate = (1 + risk_free_curve.interpolate(i) + spread_curve.interpolate(i)/10000) ** i
            duration += i * cash_flow / discount_rate

        return duration / price

    def sensitivity(self, risk_free_curve, spread_curve, libor_curve=None):
        return -self.duration(risk_free_curve, spread_curve, libor_curve)/(1+risk_free_curve.interpolate(self.maturity))

    def calculate_cash_flow(self, period, risk_free_curve, spread_curve, libor_curve):
        first_coupon = self.maturity - int(self.maturity) + 1 / self.coupon_frequency
        if self.bond_type == 'bullet':
            if self.coupon_rate_type == 'Fixed':
                return (self.coupon_rate_or_margin / self.coupon_frequency) if period < self.maturity else 1 + (self.coupon_rate_or_margin / self.coupon_frequency) 
            elif self.coupon_rate_type == 'Variable':
                return ((self.coupon_rate_or_margin/10000 + libor_curve.interpolate(period)) / self.coupon_frequency) if period < self.maturity else 1 + ((self.coupon_rate_or_margin/10000 + libor_curve.interpolate(period)) / self.coupon_frequency)
        elif self.bond_type == 'fixed annuities':
            return self.calculate_annuity()
        elif self.bond_type == 'constant amortizations':
            principal_payment = 100 / (self.maturity - first_coupon + 1)
            interest_payment = 100 * (self.maturity * self.coupon_frequency - (period - 1)) / (self.maturity * self.coupon_frequency) * (self.coupon_rate_or_margin / self.coupon_frequency)
            return principal_payment + interest_payment
        elif self.bond_type == 'equal series repayment':
            amortization=0
            first_coupon = self.maturity - int(self.maturity) + 1
            first_coupon_with_amortization = self.maturity - int(self.maturity) + 1 / self.coupon_frequency
            total_payments = int(self.maturity * self.coupon_frequency)
            series_payment = 100 / total_payments

            for i in (first_coupon + j for j in range(int(self.maturity - first_coupon) + 1)):
                outstanding_capital = 100 - amortization

                interest_payment = outstanding_capital * (self.coupon_rate_or_margin)
                total_payment = interest_payment
                for j in range(total_payments):
                    if i == first_coupon_with_amortization + j * (1 / self.coupon_frequency):
                        total_payment = interest_payment + series_payment
                        amortization += series_payment
            return total_payment
        elif self.bond_type == 'zero coupon':
            if period == self.maturity :
                return 100
            else :
                return 0

    def schedule(self, risk_free_curve, spread_curve, libor_curve=None):
        schedule = []
        first_coupon = self.maturity - int(self.maturity) + 1 / self.coupon_frequency
        if self.bond_type=="equal series repayment":
            amortization=0
            first_coupon = self.maturity - int(self.maturity) + 1
            first_coupon_with_amortization = self.maturity - int(self.maturity) + 1 / self.coupon_frequency
            total_payments = int(self.maturity * self.coupon_frequency)
            series_payment = 100 / total_payments

            for i in (first_coupon + j for j in range(int(self.maturity - first_coupon) + 1)):
                outstanding_capital = 100 - amortization

                interest_payment = outstanding_capital * (self.coupon_rate_or_margin)
                total_payment = interest_payment
                for j in range(total_payments):
                    if i == first_coupon_with_amortization + j * (1 / self.coupon_frequency):
                        total_payment = interest_payment + series_payment
                        amortization += series_payment

                schedule.append((i, total_payment))
        else :
            for i in (first_coupon + j * (1 / self.coupon_frequency) for j in range(int((self.maturity - first_coupon) * self.coupon_frequency) + 1)):
                cash_flow = self.calculate_cash_flow(i, risk_free_curve, spread_curve, libor_curve)
                schedule.append((i, cash_flow))

        return pd.DataFrame(schedule, columns=['Period', 'Cash Flow'])


########################################################################################################################

class CURVE:

    def __init__(self, name, maturities, rates):
        self.name = name
        self.maturities = maturities
        self.rates = rates

    def interpolate(self, maturity):
        # If the specified maturity is below the minimum, the closest rate is used
        if maturity <= self.maturities.iloc[0]:
            return self.rates.iloc[0]

        # If the specified maturity is above the maximum, the closest rate is used
        if maturity >= self.maturities.iloc[-1]:
            return self.rates.iloc[-1]

        # Finding the nearest maturity values
        lower_index = self.maturities[self.maturities <= maturity].idxmax()
        upper_index = self.maturities[self.maturities >= maturity].idxmin()

        lower_maturity = self.maturities.iloc[lower_index]
        upper_maturity = self.maturities.iloc[upper_index]
        lower_rate = self.rates.iloc[lower_index]
        upper_rate = self.rates.iloc[upper_index]

        # Check if upper and lower maturities are the same
        if upper_maturity == lower_maturity:
            return lower_rate

        # Linear interpolation
        return lower_rate + (maturity - lower_maturity) * (upper_rate - lower_rate) / (upper_maturity - lower_maturity)


########################################################################################################################

# Example bond
example_bond = BOND_PRICER(
    bond_type="equal series repayment",
    company="21st Century Fox America Inc",
    maturity=4.9,
    coupon_rate_type="Variable",
    coupon_rate_or_margin=0.05,
    coupon_frequency=0.5)

# Assuming LIBOR, risk-free and spread curves data are in an Excel file
excel_path = # path to 'data for python project.xlsx' 
if not os.path.exists(excel_path):
    print(f"File {excel_path} does not exist.")
    sys.exit(1)

# Read data and construct the curves
#LIBOR
LIBOR_table = pd.read_excel(excel_path, sheet_name='Libor 3M Curve')
arr_Maturity_LIBORCurve=LIBOR_table.iloc[:,0]
arr_Rate_LIBORCurve=LIBOR_table.iloc[:,1]
libor_curve=CURVE('LIBOR',arr_Maturity_LIBORCurve,arr_Rate_LIBORCurve)    #changer les prefixes arr en df (dataframes)
#Risk-Free Curve
RF_table=pd.read_excel(excel_path, sheet_name='US Yield Curve')
arr_Maturity_RfCurve=RF_table.iloc[:,0]
arr_Rate_RfCurve=RF_table.iloc[:,1]
risk_free_curve = CURVE('risk-free',arr_Maturity_RfCurve, arr_Rate_RfCurve)

Spread_table=pd.read_excel(excel_path, sheet_name='CDX_IG_Prices')
arr_Maturity_SpreadCurve = Spread_table.iloc[:,0]
# Ensure the company exists in the table
if example_bond.company in Spread_table.columns:
    # Extract the values from the identified column
    arr_Rate_SpreadCurve = Spread_table[example_bond.company]
else:
    print(f"Column for {example_bond.company} not found in Spread_table.")
    sys.exit(1)
spread_curve=CURVE(example_bond.company+' Spread',arr_Maturity_SpreadCurve, arr_Rate_SpreadCurve)

# Printing bond price
try:
    print("Bond price:", example_bond.price(risk_free_curve, spread_curve, libor_curve))
except Exception as e:
    print(f"Error while calculating bond price: {e}")

# Printing duration
try:
    print("Bond duration:", example_bond.duration(risk_free_curve, spread_curve, libor_curve))
except Exception as e:
    print(f"Error while calculating bond duration: {e}")

# Printing sensitivity
try:
    print("Bond sensitivity:", example_bond.sensitivity(risk_free_curve, spread_curve, libor_curve))
except Exception as e:
    print(f"Error while calculating bond sensitivity: {e}")

# Printing schedule
try:
    print("Bond schedule:")
    print(example_bond.schedule(risk_free_curve, spread_curve, libor_curve))
except Exception as e:
    print(f"Error while calculating bond schedule: {e}")
