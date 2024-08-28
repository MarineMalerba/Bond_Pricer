# Bond_Pricer

## Overview
The **Bond Pricer** Project is a Python-based tool designed to calculate the price, duration, sensitivity, and payment schedule of various types of bonds. The project supports different bond types, including bullet bonds, zero-coupon bonds, fixed annuities, constant amortizations, and equal series repayments. The tool also takes into account different interest rate curves like the risk-free curve, spread curve, and LIBOR curve.

## Features
+ **Bond Pricing**: Calculate the price of bonds based on the type, maturity, coupon rate, and relevant financial curves.
+ **Duration Calculation**: Compute the duration of a bond, which measures its sensitivity to interest rate changes.
+ **Sensitivity Calculation**: Determine the bond's sensitivity, which reflects the percentage change in price for a given change in interest rates.
+ **Cash Flow Schedule**: Generate a payment schedule that shows the cash flows over the life of the bond.
+ **Support for Multiple Bond Types**:
  + Bullet Bonds
  + Zero-Coupon Bonds
  + Fixed Annuities
  + Constant Amortizations
  + Equal Series Repayment Bonds

## Project Structure
+ ``BOND_PRICER Class``: Core class that handles bond pricing and related calculations.
+ ``CURVE Class``: Handles interpolation of rates over a given maturity range for interest rate curves (risk-free, spread, LIBOR).
+ Main Script: Example script that demonstrates how to use the ``BOND_PRICER`` and ``CURVE`` classes to price a bond, calculate its duration, sensitivity, and generate a cash flow schedule.
