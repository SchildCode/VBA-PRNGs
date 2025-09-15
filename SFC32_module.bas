'===============================================================
' SFC32, optimized for VBA on 64-bit Office
' Small Fast Counting (SFC) Pseudo-Random Number Generator (PRNG)
' Created by Chris Doty-Humphrey
' Ref. https://pracrand.sourceforge.net/RNG_engines.txt
' Translated to VBA by Peter.Schild@OsloMet.no 2025
' Note: On 64-bit PCs running 64-bit Office, it's fastest to use native 64-bit types [LongLong (suffix ^) for integer arithmetic, and Double (suffix #) for floating-point arithmetic]
'===============================================================

Option Explicit
Private Const BIT3^ = 2 ^ 3
Private Const BIT9^ = 2 ^ 9
Private Const BIT11^ = 2 ^ 11
Private Const BIT15^ = 2 ^ 15
Private Const BIT16^ = 2 ^ 16
Private Const BIT21^ = 2 ^ 21
Private Const MASK32^ = 2 ^ 32 - 1
Private Const PHI32^ = 2654435769^ '&H9E3779B9 Golden ratio
Private Const DIV32# = 2.3283064365387E-10 '=1/(2^32)
Private a^, b^, c^, d^ 'SFC state variables (initialized by subroutine SFC32_init)

Sub SFC32_init(Optional seed^ = 1^)
    'Initializes the SFC32 pseudorandom number generator
    'It sets 4 distinct state variables a,b,c,d, and warms up
    Dim s^
    s = seed And MASK32 'Ensures that the seed is in range [0, 2^32-1]
    a = mix32(s) 'SplitMix32 mixer/finaalizer
    b = mix32(s)
    c = mix32(s)
    d = mix32(s)
    If (a Or b Or c) = 0 Then a = a Xor &H1    'avoid all-zero corner
    For s^ = 1 To 12
        Call SFC32# '12-step warm-up (optional but recommended)
    Next
End Sub

Function SFC32#(Optional Nmax^ = 1^)
    'Generates next SFC32 pseudorandom number. Remember to first inititialize & seed with SFC32_init(seed)
    'Arguments: Nmax=1 for decimal output [0,1), and integer 1<Nmax for integer output [1,Nmax] avoiding modulo bias
    'Note: First inititlize & seed with routine SFC32_init() to set globals a,b,c,d, before using this function
    Dim t^, cL^, cR^
reject:
    t = (a + b + d) And MASK32                   't = (a + b + d) | 0
    d = (d + 1) And MASK32                       'd = d + 1 | 0;
    a = b Xor (b \ BIT9)                         'a = b ^ b >>> 9
    b = (c + ((c * BIT3) And MASK32)) And MASK32 'b = c + (c << 3) | 0
    cL = (c * BIT21) And MASK32                  'cL = c << 21
    cR = c \ BIT11                               'cR >>> 11
    c = ((cL Or cR) + t) And MASK32              'c = (cL|cR) + t | 0
    t = t And MASK32                             '(t >>> 0) // 4294967296
    'Return a Double-precision type containing either a decimal or whole number, depending on the value of argument Nmax
    If Nmax <= 1^ Then 'return decimal in range [0,1)
        SFC32 = t * DIV32 'Output unsigned double (32-bit) in the range [0,1); it can potentially equal 0, but is always less than 1.
    Else 'return integer in range [1,Nmax] avoiding modulo bias by rejection-sampling. Nmax must be an integer in the range [2, 2*32]
        'Note: I also tried Algorithm 5 by Daniel Lemire (https://arxiv.org/abs/1805.10941), but it was actually NOT systematically faster then this one (\ is replaced by AND, which is just as slow on VBA).
        SFC32 = 1# + t \ (MASK32 \ Nmax) 'Note: It is possible to precumpute the static denominator (MASK32 \ Nmax), but the speedup is only 3%
        If Nmax < SFC32 Then GoTo reject
    End If
End Function

Private Function mix32^(ByRef s^)
    'Helper function for SFC32
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

'====================================== Alt.2: encapsulated single-subroutine for SFC32 with integrated initialization

Function xSFC32#(Optional Nmax^ = 1^)
    'Self-contained extra-Small Fast Counter (xSFC) random number generator, with self-contained initialization and hard-coded seed
    'Passes DieHard and TestU01 BigCrush tests
    'Arguments: Nmax=1 for decimal output [0,1), and integer 1<Nmax for integer output [1,Nmax] avoiding modulo bias
    Const BIT3^ = 2 ^ 3
    Const BIT9^ = 2 ^ 9
    Const BIT11^ = 2 ^ 11
    Const BIT21^ = 2 ^ 21
    Const MASK32^ = 2 ^ 32 - 1
    Const DIV32# = 2.3283064365387E-10 '=1/(2^32)
    Dim tt^, cL^, cR^
    Static aa^, bb^, cc^, dd^, init^
    If Not init Then init = True: aa = 2754059347^: bb = 288685983^: cc = 479643687^: dd = 2575288963^ 'This combination from SplitMix32 gives relatively uniform distribution for the first 10000 numbers
'---
reject:
    tt = (aa + bb + dd) And MASK32                  't = (a + b + d) | 0
    dd = (dd + 1) And MASK32                        'd = d + 1 | 0;
    aa = bb Xor (bb \ BIT9)                         'a = b ^ b >>> 9
    bb = (cc + ((cc * BIT3) And MASK32)) And MASK32 'b = c + (c << 3) | 0
    cL = (cc * BIT21) And MASK32                    'cL = c << 21
    cR = cc \ BIT11                                 'cR >>> 11
    cc = ((cL Or cR) + tt) And MASK32               'c = (cL|cR) + t | 0
    tt = tt And MASK32                              '(t >>> 0) // 4294967296
    'Return a Double-precision type containing either a decimal or whole number, depending on the value of argument Nmax
    If Nmax <= 1^ Then 'return decimal in range [0,1)
        xSFC32 = tt * DIV32 'Output unsigned double (32-bit) in the range [0,1); it can potentially equal 0, but is always less than 1.
    Else 'return integer in range [1,Nmax] avoiding modulo bias by rejection-sampling. Nmax must be an integer in the range [2, 2*32]
        xSFC32 = 1# + tt \ (MASK32 \ Nmax) 'Note: It is possible to precumpute the static denominator (MASK32 \ Nmax), but the speedup is only 3%
        If Nmax < xSFC32 Then GoTo reject
    End If
End Function

'====================================== The code below is just for testing, and can be deleted ======================================

Private Sub test_SFC()
    'This is a throw-away subroutine just to check that PCG32 is working
    Dim i&
    Call SFC32_init(12345^) 'initialise, and seed with a LongLong integer [0, 2^32-1]
    For i = 1 To 10
        Debug.Print i, SFC32(), SFC32(1000^) 'SFC32() returns a decimal U[0,1), and SFC32(1000) returns an integer in the range [1,1000] without modulo bias.
    Next
End Sub

Function SFC32_U32^()
    'Core algorithm of the SFC32 pseudorandom number generator
    'This is the "Bryc" variant of SFC32, not the original SFC32 by Chris Doty-Humphrey
    'State: 4×32-bit (128 bits total), period about 2^128.
    'While Doty-Humphrey’s original SFC32 uses multiply add, xor, and rotate, Bryc’s variant drops the multiply
    'Remember to first inititialize & seed with SFC32_init(seed)
    'Ref. https://github.com/bryc/code/issues/11
    'Ref. https://github.com/bryc/code/blob/master/jshash/PRNGs.md  (under heading 'sfc32')
    'Function: Returns an unsigned 32-bit integer in a LongLong, useful for testing on TestU01 or ENT. This is because the 32-bit Long-type is signed
    Dim t^, cL^, cR^
'---
    t = (a + b + d) And MASK32                   't = (a + b + d) | 0
    d = (d + 1) And MASK32                       'd = d + 1 | 0;
    a = b Xor (b \ BIT9)                         'a = b ^ b >>> 9
    b = (c + ((c * BIT3) And MASK32)) And MASK32 'b = c + (c << 3) | 0
    cL = (c * BIT21) And MASK32                  'cL = c << 21
    cR = c \ BIT11                               'cR >>> 11
    c = ((cL Or cR) + t) And MASK32              'c = (cL|cR) + t | 0
    SFC32_U32 = t And MASK32                     '(t >>> 0) // 4294967296
End Function

Function xSFC32_U32^()
    'extra-Small Fast Counter (xSFC) random number generator, with self-contained initialization (same as function xSFC32)
    'Function: Returns an unsigned 32-bit integer in a LongLong, useful for testing on TestU01 or ENT. This is because the 32-bit Long-type is signed
    Dim tt^, cL^, cR^
    Static aa^, bb^, cc^, dd^, init^
    If Not init Then init = True: aa = 2754059347^: bb = 288685983^: cc = 479643687^: dd = 2575288963^ 'This combination from SplitMix32 gives relatively uniform distribution for the first 10000 numbers
'---
    tt = (aa + bb + dd) And MASK32                  't = (a + b + d) | 0
    dd = (dd + 1) And MASK32                        'd = d + 1 | 0;
    aa = bb Xor (bb \ BIT9)                         'a = b ^ b >>> 9
    bb = (cc + ((cc * BIT3) And MASK32)) And MASK32 'b = c + (c << 3) | 0
    cL = (cc * BIT21) And MASK32                    'cL = c << 21
    cR = cc \ BIT11                                 'cR >>> 11
    cc = ((cL Or cR) + tt) And MASK32               'c = (cL|cR) + t | 0
    xSFC32_U32 = tt And MASK32                      '(t >>> 0) // 4294967296
End Function

