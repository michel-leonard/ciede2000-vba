' This function written in VBA is not affiliated with the CIE (International Commission on Illumination),
' and is released into the public domain. It is provided "as is" without any warranty, express or implied.

' The classic CIE ΔE2000 implementation, which operates on two L*a*b* colors, and returns their difference.
' "l" ranges from 0 to 100, while "a" and "b" are unbounded and commonly clamped to the range of -128 to 127.
Public Function ciede2000(l1 As Double, a1 As Double, b1 As Double, l2 As Double, a2 As Double, b2 As Double, Optional kl As Double = 1.0, Optional kc As Double = 1.0, Optional kh As Double = 1.0, Optional canonical As Boolean = False) As Double
	' Working in VBA with the CIEDE2000 color-difference formula.
	' kl, kc, kh are parametric factors to be adjusted according to
	' different viewing parameters such as textures, backgrounds...
	Const M_PI = 3.14159265358979323846264338328
	Dim n As Double, c1 As Double, c2 As Double, h1 As Double, h2 As Double
	Dim h_mean As Double, h_delta As Double, p As Double, r_t As Double
	Dim l As Double, t As Double, h As Double, c As Double
	n = (Sqr(a1 * a1 + b1 * b1) + Sqr(a2 * a2 + b2 * b2)) * 0.5
	n = n * n * n * n * n * n * n
	' A factor involving chroma raised to the power of 7 designed to make
	' the influence of chroma on the total color difference more accurate.
	n = 1.0 + 0.5 * (1.0 - Sqr(n / (n + 6103515625.0)))
	' Application of the chroma correction factor.
	c1 = Sqr(a1 * a1 * n * n + b1 * b1)
	c2 = Sqr(a2 * a2 * n * n + b2 * b2)
	' Using 14 lines to simulate atan2, as VBA does not have this built-in.
	If 0.0 < a1 Then
		h1 = Atn(b1 / (a1 * n)) - (b1 < 0.0) * 2.0 * M_PI
	ElseIf a1 < 0.0 Then
		h1 = Atn(b1 / (a1 * n)) + M_PI
	Else
		h1 = M_PI + ((0.0 < b1) - (b1 < 0.0)) * 0.5 * M_PI
	End If
	If 0.0 < a2 Then
		h2 = Atn(b2 / (a2 * n)) - (b2 < 0.0) * 2.0 * M_PI
	ElseIf a2 < 0.0 Then
		h2 = Atn(b2 / (a2 * n)) + M_PI
	Else
		h2 = M_PI + ((0.0 < b2) - (b2 < 0.0)) * 0.5 * M_PI
	End If
	' The atan2 polyfill (customized) is complete.
	' When the hue angles lie in different quadrants, the straightforward
	' average can produce a mean that incorrectly suggests a hue angle in
	' the wrong quadrant, the next lines handle this issue.
	h_mean = (h1 + h2) * 0.5
	h_delta = (h2 - h1) * 0.5
	' The part where most programmers get it wrong.
	If M_PI + 1E-14 < Abs(h2 - h1) Then
		h_delta = h_delta + M_PI
		If canonical And M_PI + 1E-14 < h_mean Then
			' Sharma’s implementation, OpenJDK, ...
			h_mean = h_mean - M_PI
		else
			' Lindbloom’s implementation, Netflix’s VMAF, ...
			h_mean = h_mean + M_PI
		End If
	End If
	p = 36.0 * h_mean - 55.0 * M_PI
	n = (c1 + c2) * 0.5
	n = n * n * n * n * n * n * n
	' The hue rotation correction term is designed to account for the
	' non-linear behavior of hue differences in the blue region.
	r_t = -2.0 * Sqr(n / (n + 6103515625.0)) _
			* Sin(M_PI / 3.0 * Exp(p * p / (-25.0 * M_PI * M_PI)))
	n = (l1 + l2) * 0.5
	n = (n - 50.0) * (n - 50.0)
	' Lightness.
	l = (l2 - l1) / (kl * (1.0 + 0.015 * n / Sqr(20.0 + n)))
	' These coefficients adjust the impact of different harmonic
	' components on the hue difference calculation.
	t = 1.0	- 0.17 * Sin(h_mean + M_PI / 3.0) _
			+ 0.24 * Sin(2.0 * h_mean + M_PI * 0.5) _
			+ 0.32 * Sin(3.0 * h_mean + 8.0 * M_PI / 15.0) _
			- 0.2 * Sin(4.0 * h_mean + 3.0 * M_PI / 20.0)
	n = c1 + c2
	' Hue.
	h = 2.0 * Sqr(c1 * c2) * Sin(h_delta) / (kh * (1.0 + 0.0075 * n * t))
	' Chroma.
	c = (c2 - c1) / (kc * (1.0 + 0.0225 * n))
	' The result reflects the actual geometric distance in the color space, given a tolerance of 3.6e-13.
	ciede2000 = Sqr(l * l + h * h + c * c + c * h * r_t)
End Function

' If you remove the constant 1E-14, the code will continue to work, but CIEDE2000
' interoperability between all programming languages will no longer be guaranteed.

' Source code tested by Michel LEONARD
' Website: ciede2000.pages-perso.free.fr
