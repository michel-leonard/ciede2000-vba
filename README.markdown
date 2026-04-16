Browse : [Ruby](https://github.com/michel-leonard/ciede2000-ruby) · [Rust](https://github.com/michel-leonard/ciede2000-rust) · [SQL](https://github.com/michel-leonard/ciede2000-sql) · [Swift](https://github.com/michel-leonard/ciede2000-swift) · [TypeScript](https://github.com/michel-leonard/ciede2000-typescript) · **VBA** · [Wolfram Language](https://github.com/michel-leonard/ciede2000-wolfram-language) · [AWK](https://github.com/michel-leonard/ciede2000-awk) · [BC](https://github.com/michel-leonard/ciede2000-basic-calculator) · [C#](https://github.com/michel-leonard/ciede2000-csharp) · [C++](https://github.com/michel-leonard/ciede2000-cpp)

# CIEDE2000 color difference formula in VBA

This page presents the CIEDE2000 color difference, implemented in the VBA programming language.

![Logo for CIEDE2000 in Visual Basic for Applications](https://raw.githubusercontent.com/michel-leonard/ciede2000-color-matching/refs/heads/main/docs/assets/images/logo.jpg)

## About

Here you’ll find the first rigorously correct implementation of CIEDE2000 that doesn’t use any conversion between degrees and radians. Set parameter `canonical` to obtain results in line with your existing pipeline.

`canonical`|The algorithm operates...|
|:--:|-|
`False`|in accordance with the CIEDE2000 values currently used by many industry players|
`True`|in accordance with the CIEDE2000 values provided by [this](https://hajim.rochester.edu/ece/sites/gsharma/ciede2000/) academic MATLAB function|

## Our CIEDE2000 offer

This production-ready file, released in 2026, contain the CIEDE2000 algorithm.

Source File|Type|Bits|Purpose|Advantage|
|:--:|:--:|:--:|:--:|:--:|
[ciede2000.bas](./ciede2000.bas)|`Double`|64|General|Interoperability|

> A [native Microsoft Excel formula for CIEDE2000](https://github.com/michel-leonard/ciede2000-excel) is also available if portability is your priority.

### Software Versions

All versions of VBA from Excel 97 to Excel 365, Windows and Mac.

### Example Usage

We calculate the CIEDE2000 distance between two colors, first without and then with parametric factors.

```vba
Sub Main()

	' Example of two L*a*b* colors
	Dim l1 As Double: l1 = 76.8
	Dim a1 As Double: a1 = 103.9
	Dim b1 As Double: b1 = 6.6

	Dim l2 As Double: l2 = 73.6
	Dim a2 As Double: a2 = 116.1
	Dim b2 As Double: b2 = -3.9

	Dim deltaE As Double

	deltaE = ciede2000(l1, a1, b1, l2, a2, b2)
	Debug.Print "CIEDE2000 = "; deltaE
	' ΔE2000 = 4.575648907164364

	' Example of parametric factors used in the textile industry
	Dim kl As Double: kl = 2.0
	Dim kc As Double: kc = 1.0
	Dim kh As Double: kh = 1.0

	' Perform a CIEDE2000 calculation compliant with that of Gaurav Sharma
	Dim canonical As Boolean: canonical = True

	deltaE = ciede2000(l1, a1, b1, l2, a2, b2, kl, kc, kh, canonical)
	Debug.Print "CIEDE2000 = "; deltaE
	' ΔE2000 = 4.1058032084100855

End Sub
```

### Test Results

LEONARD’s tests are based on well-chosen L\*a\*b\* colors, with various parametric factors `kL`, `kC` and `kH`.

```
CIEDE2000 Verification Summary :
          Compliance : [ ] CANONICAL [X] SIMPLIFIED
  First Checked Line : 20.0,0.05,-30.0,30.0,0.0,128.0,1.0,1.0,1.0,53.41746217641311
           Precision : 12 decimal digits
           Successes : 10000000
               Error : 0
            Duration : 184.56 seconds
     Average Delta E : 67.13
   Average Deviation : 1e-14
   Maximum Deviation : 3.1e-13
```

```
CIEDE2000 Verification Summary :
          Compliance : [X] CANONICAL [ ] SIMPLIFIED
  First Checked Line : 20.0,0.05,-30.0,30.0,0.0,128.0,1.0,1.0,1.0,53.41765416511742
           Precision : 12 decimal digits
           Successes : 10000000
               Error : 0
            Duration : 181.35 seconds
     Average Delta E : 67.13
   Average Deviation : 1e-14
   Maximum Deviation : 3.1e-13
```

## Public Domain Licence

You are free to use these files, even for commercial purposes.
