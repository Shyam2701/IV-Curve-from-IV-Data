import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy.optimize import curve_fit

# Load the Excel file
file_path = 'D:/Shyampy/IV_data(Lighted).xlsx'
data = pd.read_excel(file_path)

voltage = data['Voltage(V)']
current = data['Current(A)']

data["Power(W)"]= voltage * current
max_power_index=data["Power(W)"].idxmax()
max_voltage=data.loc[max_power_index, 'Voltage(V)']
max_current=data.loc[max_power_index, 'Current(A)']
max_power=data.loc[max_power_index, 'Power(W)']


# Plot the IV-curve
plt.figure(figsize=(10, 6))
plt.plot(voltage, current, linestyle='-', color='b')
plt.scatter(max_voltage,max_current, marker='o', color='r')
plt.title('IV-Curve')
plt.xlabel('Voltage (V)')
plt.ylabel('Current (A)')
plt.grid(True)
plt.show()

# Define a linear function for fitting
def linear_func(x, a, b):
    return a * x + b

# Calculate Rs (slope near Voc)
# We'll use data near Voc for this calculation
# Voc is when current is near 0 (open-circuit)
Voc_region = data[voltage > 0.6 * max(voltage)]
popt_rs, _ = curve_fit(linear_func, Voc_region['Voltage(V)'], Voc_region['Current(A)'])
Rs = -1 / popt_rs[0]
print(popt_rs)

# Calculate Rsh (slope near Isc)
# Isc is when voltage is near 0 (short-circuit)
Isc_region = data[voltage < 0.5 * max(voltage)]
popt_rsh, _ = curve_fit(linear_func, Isc_region['Voltage(V)'], Isc_region['Current(A)'])
Rsh = -1 / popt_rsh[0]

print("Vmp: ", round(max_voltage,2), "V")
print("Imp: ", round(max_current,2), "A")
print("Pmax.: ", round(max_power,2), "W")
print("Rsh: ", round(Rsh,2), "Ohm")
print("Rs: ", round(Rs,2), "Ohm")
