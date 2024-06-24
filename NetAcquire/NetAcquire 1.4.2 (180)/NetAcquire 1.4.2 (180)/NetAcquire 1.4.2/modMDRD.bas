Attribute VB_Name = "modMDRD"
Option Explicit

'MDRD equation
'eGFR = 32788 * ((serum creat)^-1.158) * Age^-0.203
'( * 1.180 if black )
'( * 0.742 if female )
'units=mls/min/1.73m^2

'So for a white male of 70 years old, where the measured creatinine on an Abbott system was 90 umol/L
'GFR (mL/min/1.73 m2) = 175 x [serum creatinine x 0.011312]-1.154 x [age]-0.203
'175 x [90x 0.011312]-1.154 x [70]-0.203 is the non-adjusted form,
'but I suggest you use the following instead
'175 x [((90-intercept)/slope)x 0.011312]-1.154 x [70]-0.203
'and from the look-up table below we see that the Abbott intercept and slope are 13.21 and 0.940
'respectively, so we get
'175 x [(90-13.21)/0.940))x 0.011312]-1.154 x [70]-0.203
'which is an eGFR of 81
'175 x [(sCreat-13.21)/0.940))x 0.011312]^-1.154 x [Age]^-0.203


