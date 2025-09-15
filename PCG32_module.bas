'===============================================================
' PCG32 (XSH-RR), optimized for VBA on 64-bit Office
' Permuted Congruential Generator (PCG) pseudo-random number generator (PRNG)
' Ref. https://en.wikipedia.org/wiki/Permuted_congruential_generator
' Ref. https://www.pcg-random.org
' Translated to VBA by Peter.Schild@OsloMet.no 2025
' Note: On 64-bit PCs running 64-bit Office, it's fastest to use native 64-bit types [LongLong (suffix ^) for integer arithmetic, and Double (suffix #) for floating-point arithmetic]
'===============================================================

Option Explicit
Private Const PCG_M_hi^ = &H5851F42D '0x5851F42D4C957F2D >> 32
Private Const PCG_M_lo^ = &H4C957F2D 'low 32 bits
Private Const BIT5^ = 2 ^ 5
Private Const BIT14^ = 2 ^ 14
Private Const BIT15^ = 2 ^ 15
Private Const BIT16^ = 2 ^ 16
Private Const BIT18^ = 2 ^ 18
Private Const BIT27^ = 2 ^ 27
Private Const BIT31^ = 2 ^ 31
Private Const BIT32^ = 2 ^ 32
Private Const MASK16^ = BIT16 - 1 '2^16 -1
Private Const MASK18^ = BIT18 - 1 '2^18 -1
Private Const MASK27^ = BIT27 - 1 '2^27 -1
Private Const MASK32^ = BIT32 - 1 '2^32 -1
Private Const DIV32# = 2.3283064365387E-10 '=1/(2^32)
Private Const PHI32^ = 2654435769^ '&H9E3779B9 Golden ratio
Private s_hi^, s_lo^, i_hi^, i_lo^ '64-bit state and 64-bit increment (odd), hi:lo (each 0..2^32-1)
Private SH^(0 To 32) 'SH(k) = 2^k, tiny precalculated table for variable rotate

Public Sub PCG32_init(Optional seed^ = 1^)
    'Initializes the PCG32 pseudorandom number generator, seeds the 64-bit state and 64-bit increment
    'Precalculates 2^k table, and builds the 4 distinct
    Dim w1^, w2^, w3^, w4^
    Dim s^, sum^, carry^
    Dim k&
'---
    SH(0) = 1&
    For k = 1& To 32&
        SH(k) = SH(k - 1&) + SH(k - 1&) 'Precalculate table of 2^k by exact doubling in integer
    Next
'---
    '1) Four distinct 32-bit words from master seed via SplitMix32
    s = seed And MASK32 'Ensures that the seed is in range [0, 2^32-1]
    w1 = mix32(s) 'SplitMix32 mixer/finaalizer
    w2 = mix32(s)
    w3 = mix32(s)
    w4 = mix32(s)

    '2) Build inc = (seq<<1)|1 with 64-bit (hi:lo). Use w1 for the high 32, w2 for the low 32 of seq; then double + make odd.
    i_hi = (w1 + w1) And MASK32
    i_lo = ((w2 + w2) And MASK32) Or 1^
    If (w2 And BIT31) <> 0^ Then i_hi = (i_hi + 1^) And MASK32 'carry from low doubling into high

    '3) Build state0 = (w3:w4) and do O’Neill’s seeding: state=0; advance(); state+=state0; advance()
    s_hi = 0^: s_lo = 0^
    PCG32_AdvanceOnce s_hi, s_lo
    'state += state0 (mod 2^64) using 32-bit limbs
    sum = s_lo + w4
    carry = IIf(BIT32 <= sum, 1^, 0^)
    s_lo = sum And MASK32
    s_hi = (s_hi + w3 + carry) And MASK32
    PCG32_AdvanceOnce s_hi, s_lo

    'Debug.Print "Initialized PCG32 with state variables", s_hi, s_lo, i_hi, i_lo
End Sub

Public Function PCG32#(Optional Nmax^ = 1^)
    'Generates next PCG32 (XSH-RR) pseudorandom number
    'Arguments: Nmax=1 for decimal output [0,1), and integer 1<Nmax for integer output [1,Nmax] avoiding modulo bias
    'Note: First inititlize & seed with routine PCG32_init() to set globals s_hi, s_lo, i_hi, i_lo, and SH(), before using this function
    Dim old_hi^, old_lo^, rHi^, rLo^, xhi^, xlo^, x^, u^
    Dim rot&
'---
reject:
    old_hi = s_hi: old_lo = s_lo
    PCG32_AdvanceOnce s_hi, s_lo 'state = state * M + inc
    rHi = old_hi \ BIT18
    rLo = (old_lo \ BIT18) Or ((old_hi And MASK18) * BIT14)
    xhi = (rHi Xor old_hi) And MASK32
    xlo = (rLo Xor old_lo) And MASK32
    x = ((xlo \ BIT27) Or ((xhi And MASK27) * BIT5)) And MASK32
    rot = CLng(old_hi \ BIT27) And 31& ' rot = old >> 59 == old_hi >> 27
    rot = rot And 31&
    If rot = 0& Then u = x And MASK32 Else u = ((x \ SH(rot)) Or ((x * SH(32 - rot)) And MASK32)) And MASK32 'where u[0,2^32-1]
    'Return a Double-precision type containing either a decimal or whole number, depending on the value of argument Nmax
    If Nmax <= 1^ Then 'return decimal in range [0,1)
        PCG32 = u * DIV32 'Output unsigned double (32-bit) in the range [0,1); it can potentially equal 0, but is always less than 1.
    Else 'return integer in range [1,Nmax] avoiding modulo bias by rejection-sampling. Nmax must be an integer in the range [2, 2*32]
        'Note: I also tried Algorithm 5 by Daniel Lemire (https://arxiv.org/abs/1805.10941), but it was actually NOT systematically faster then this one (\ is replaced by AND, which is just as slow on VBA).
        PCG32 = 1# + u \ (MASK32 \ Nmax) 'Note: It is possible to precumpute the static denominator (MASK32 \ Nmax), but the speedup is only 3%
        If Nmax < PCG32 Then GoTo reject
    End If
End Function

Private Function mix32^(ByRef s^)
    'Helper function for PCG32.
    'SplitMix32 algorithm with Marasglai's Weyl increment followed by a xmxmx-mixer with Hash-Prospector 'optimized' constants discovered by TheIronBorn in 2022 (https://github.com/skeeto/hash-prospector/issues/19#issuecomment-1120105785)
    'This 'optimized' version is also recommended by Bryc: https://github.com/bryc/code/blob/master/jshash/PRNGs.md
    'Returns incremented state s^ [0, 2^32) and random integer mix32^ [0, 2^32), both LongLong-type variables
    s = (s + PHI32) And MASK32 'Weyl increment, typically 0x9E3779B9 from Marsaglia
    s = s Xor (s \ BIT16) 'XOR shift
    s = (s * &H21F0AAAD) And MASK32 'Multiply &H21F0AAAD
    s = s Xor (s \ BIT15) 'XOR shift
    s = (s * &H735A2D97) And MASK32 'Multiply &H735A2D97
    s = s Xor (s \ BIT15) 'XOR shift
    mix32 = s
End Function

Private Sub PCG32_AdvanceOnce(ByRef a_hi^, ByRef a_lo^)
    'Helper function for PCG32 PRNG
    'Function: One LCG step: state = state * M + inc (mod 2^64)
    Dim p0_hi^, p0_lo^, s1_lo^, tmp_hi^, s2_lo^, sLo^, c^
'---
    Mul32x32 a_lo, PCG_M_lo, p0_hi, p0_lo
    Mul32x32 a_lo, PCG_M_hi, tmp_hi, s1_lo
    Mul32x32 a_hi, PCG_M_lo, tmp_hi, s2_lo
    sLo = (s1_lo + s2_lo) And MASK32 'only low 32 bits of (s1_lo + s2_lo) survive after <<32; tmp_hi is ignored

    a_lo = p0_lo + i_lo
    If BIT32 <= a_lo Then a_lo = a_lo - BIT32: c = 1^ Else c = 0^ 'it's sufficient to wrap once
    a_hi = (p0_hi + i_hi + c + sLo) And MASK32
End Sub

Private Sub Mul32x32(ByVal a^, ByVal b^, ByRef hi^, ByRef lo^)
    'Helper function for PCG32 PRNG
    'Function: Multiply 32×32 bits and returns 64 bits as hi/lo limbs, exact. This routine is needed because VBA does not have unsigned 64-bit intengers for safe 32*32 multiplication; LongLong type is signed 64-bit integer
    Dim a0^, a1^, b0^, b1^, c0^, c1^, c2^, lo_tmp^, carry^
'---
    a0 = a And MASK16: a1 = (a \ BIT16) And MASK16
    b0 = b And MASK16: b1 = (b \ BIT16) And MASK16

    c0 = a0 * b0           '< 2^32
    c1 = a0 * b1 + a1 * b0 '< 2^33
    c2 = a1 * b1           '< 2^32

    lo_tmp = c0 + ((c1 And MASK16) * BIT16) 'lo = c0 + ((c1 & 0xFFFF) << 16)
    If BIT32 <= lo_tmp Then lo = lo_tmp - BIT32: carry = 1^ Else lo = lo_tmp: carry = 0^ '...(mod 2^32), track carry
    hi = (c2 + (c1 \ BIT16) + carry) And MASK32 'hi = c2 + (c1 >> 16) + carry  (mod 2^32)
End Sub

'====================================== The code below is just for testing, and can be deleted ======================================

Private Sub test_PCG()
    'This is a throw-away subroutine just to check that PCG32 is working
    Dim i&
    Call PCG32_init(12345^) 'initialise, and seed with a LongLong integer [0, 2^32-1]
    For i = 1 To 10
        Debug.Print i, PCG32(), PCG32(1000^) 'PCG32(1) return a decimal U[0,1), and PCG32(1000) returns an integer in the range [1,1000] without modulo bias.
    Next
End Sub

Function PCG32_U32^()
    'Core algorithm of the PCG32 pseudorandom number generator
    'Function: Returns an unsigned 32-bit integer in a LongLong, useful for testing on TestU01 or ENT. This is because the 32-bit Long-type is signed
    Dim old_hi^, old_lo^, rHi^, rLo^, xhi^, xlo^, x^, u^
    Dim rot&
'---
    old_hi = s_hi: old_lo = s_lo
    PCG32_AdvanceOnce s_hi, s_lo 'state = state * M + inc
    rHi = old_hi \ BIT18
    rLo = (old_lo \ BIT18) Or ((old_hi And MASK18) * BIT14)
    xhi = (rHi Xor old_hi) And MASK32
    xlo = (rLo Xor old_lo) And MASK32
    x = ((xlo \ BIT27) Or ((xhi And MASK27) * BIT5)) And MASK32
    rot = CLng(old_hi \ BIT27) And 31& ' rot = old >> 59 == old_hi >> 27
    rot = rot And 31&
    If rot = 0& Then u = x And MASK32 Else u = ((x \ SH(rot)) Or ((x * SH(32 - rot)) And MASK32)) And MASK32 'where u[0,2^32-1]
    PCG32_U32 = u
End Function
