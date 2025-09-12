'===============================================================
' This module contains other outdated or underperforming PRNGs for comparison with PCG32 and PFC32
' The versions in this module are all configured to output an unsigned 32-bit integer in a LongLong, useful for testing on TestU01 or ENT. This is because the 32-bit Long-type is signe
' Translated to VBA by Peter.Schild@OsloMet.no 2014-2025
'
' Overview of the pseudorandom number generators in this file, listed in order of complexity:
' rnd - VBA native function rnd()
' RAND - Excel Worksheet function RAND()
' RANDARRAY - Excel Worksheet function RANDARRAY()
' fMix32 - Canonical variant of SplitMix32 with the Murmur3 fmix32 xmxmx-mixer/finisher
' mix32 - Optimized variant of SplitMix32 with Hash-Prospector 'optimized' constants discovered by TheIronBorn in 2022
' LIFB4 - 4-lagged Fibonacci generator
' MTran - Mersenne Twister
' SFC32 - Canonical 32-bit version of Small Fast Counting (SFC) Pseudo-Random Number Generator. The code is in module SFC32_module
' Byrc32 - Bryc variant of SFC32. The code is in module SFC32_module
' PCG32 - Permuted Congruential Generator (PCG) pseudo-random number generator. The code is in module PCG32_module
'===============================================================

Option Explicit
Private Const BIT13^ = 2 ^ 15
Private Const BIT15^ = 2 ^ 15
Private Const BIT16^ = 2 ^ 16
Private Const BIT32^ = 2 ^ 32
Private Const PHI32^ = 2654435769^ '&H9E3779B9 Golden ratio
Private Const MASK32^ = BIT32 - 1  '2^32 - 1 = &H00000000FFFFFFFF

'<========= Mersenne Twister (Mtran) constants =========
Const MTn& = 624
Const MTm& = 397
Const MATRIX_A& = &H9908B0DF     '/* constant vector a */
Const UPPER_MASK& = &H80000000   '/* most significant w-r bits */
Const LOWER_MASK& = &H7FFFFFFF   '/* least significant r bits */
Const kDiffMN& = MTm& - MTn&
Const Nuplim& = MTn& - 1
Const Muplim& = MTm& - 1
Const Nplus1& = MTn& + 1
Const NuplimLess1& = Nuplim - 1
Const NuplimLessM& = Nuplim - MTm&
Const k2_31# = 2 ^ 31    '2^31   ==  2147483648 == 80000000
Const k2_31Neg# = -k2_31    '-2^31  == -2147483648 == 80000000
Const k2_31b# = k2_31 - 1#     '2^31-1 ==  2147483647 == 7FFFFFFF
Const k2_32# = 2 ^ 32     '2^32   ==  4294967296 == 0
Const k2_32b# = k2_32 - 1#     '2^32-1 ==  4294967295 == FFFFFFFF == -1
Const kMT_1# = 1# / k2_32b
'========= end of MTran constants =========>


Function rnd_U32^()
    'Excel native rnd() function
    rnd_U32 = CLngLng(Int(Rnd() * CDbl(BIT32))) 'Generate random double U[0,1) and scale to D32, then store in a LongLong
End Function

Function RAND_U32^()
    'Returns Excel RAND worksheet function. Extraordinarily slow!
    Dim R As Variant
'---
    R = Evaluate("=RAND()")
    RAND_U32 = CLngLng(Int(CDbl(R) * CDbl(BIT32))) ''Generate random double U[0,1) and scale to D32, then store in a LongLong
End Function

Function RANDARRAY_U32^()
    'Returns Excel RANDARRAY worksheet function.
    'speeds up by using static array, filling once every CHUNK values.
    Const N& = 1000& 'Generate a chunk of random numbers at a time, much faster than Evaluate("=RAND()")
    Static i&
    Static R() As Variant 'store array of CHUNK random numbers. Reduces the number of slow COM calls to WorksheetFunction.RandArray
'---
    i = i + 1
    If N < i Then i = 1
    If i = 1 Then
        R = Application.WorksheetFunction.RandArray(N, 1, 0, BIT32 - 1, True) 'fill R with with N integers [0,2^32-1] without modulo bias
    End If
    RANDARRAY_U32 = CLngLng(R(i, 1)) 'convert to LongLong in range [0, 2^32-1]
End Function

Function fmix32_U32^(ByRef s^)
    'fmix32: A compact 32-bit mixer with good avalanche effect
    'Canonical variant of SplitMix32 algorithm with Marasglai's Weyl increment followed by the Murmur3 fmix32 xmxmx-mixer/finisher
    'Returns incremented state s^ [0, 2^32) and new random integer fmix32^ [0, 2^32), both LongLong-type variables
    'The seminal source for the SplitMix32 algorithm is the paper Steele, Lea & Flood (2014) "Fast Splittable Pseudorandom Number Generators", https://doi.org/10.1145/2714064.2660195
    
    s = (s + PHI32) And MASK32 'Weyl increment, typically 0x9E3779B9 from Marsaglia
    s = s Xor (s \ BIT16) 'XOR shift
    s = (s * &H85EBCA6B) And MASK32 'Multiply
    s = s Xor (s \ BIT13) 'XOR shift
    s = (s * &HC2B2AE35) And MASK32 'Multiply
    s = s Xor (s \ BIT16) 'XOR shift
    fmix32_U32 = s
End Function

Function mix32_U32^(ByRef s^)
    'mix32: A compact 32-bit mixer with good avalanche effect
    'SplitMix32 algorithm with Marasglai's Weyl increment followed by a xmxmx-mixer with Hash-Prospector 'optimized' constants discovered by TheIronBorn in 2022 (https://github.com/skeeto/hash-prospector/issues/19#issuecomment-1120105785)
    'Returns incremented state s^ [0, 2^32) and new random integer mix32^ [0, 2^32), both LongLong-type variables
    'This specific VBA translation passes the ENT test (entropy 7.999999) and DieHard test-battery, and TestU01 SmallCrush, but fails TestU01 Crush because of limited period of 2^32
    
    s = (s + PHI32) And MASK32 'Weyl increment, typically 0x9E3779B9 from Marsaglia
    s = s Xor (s \ BIT16) 'XOR shift
    s = (s * &H21F0AAAD) And MASK32 'Multiply &H21F0AAAD
    s = s Xor (s \ BIT15) 'XOR shift
    s = (s * &H735A2D97) And MASK32 'Multiply &H735A2D97
    s = s Xor (s \ BIT15) 'XOR shift
    mix32_U32 = s
End Function

Function LFIB4_U32^(Optional seed^ = 0^)
    'LFIB4 is a 4-lagged Fibonacci generator: x(n)=x(n-r) op x(n-s), with the x's in a finite set over which there is a
    'binary operation op, such as +,- on integers mod 2^32, * on odd such integers, exclusive-or(xor) on binary vectors.
    'This RNG was proposed by Marsaglia in 1999. It passes all Diehard and Big Crush tests. He claimed that it is as good as Mersenne Twister
    'Its period is 2^31*(2^256-1), about 2^287
    'Paper about its performance: https://www.sciencedirect.com/science/article/pii/S0743731519304885
    'Converted to Visual Basic by Peter G. Schild, HiOA, 2014-2025
    'P.G. Schild tested it against DieHard (http://www.stat.fsu.edu/pub/diehard/) and ENT (http://www.fourmilab.ch/random/), using the online service at http://www.cacert.at/random/
    'ENT-test Entropy = 7.999999 bits/byte. https://www.fourmilab.ch/random/
    'Test results are documented on webpage http://www.cacert.at/cgi-bin/rngresults, see row "LFIB4b (VBA port of LFIB4) corrected v.2" dated 2025-08
    Const BIT1& = 1&
    Const BIT8& = 256&
    Const BIT32# = 4294967296# '=2^32 in Double precision
    Const DIV32# = 2.3283064365387E-10 '=1/(2^32)
    Dim s^
    Dim cast#, divi#
    Static tt#(0 To BIT8) 'table of unsigned 32-bit integers, only indices 0-255 are used. Values in array tt() have range [0, 2^32), i.e. 0 <= tt() < 2^32
    Static b1&, b2&, b3&, b4&
    Static initialized As Boolean
'---
    'Initialize if necessary
    If 0 < seed Or Not initialized Then
        s = seed
        For b1 = 0 To BIT8 - 1 'only 0-255 are used
            tt(b1) = mix32_U32(s) 'Initialize with SplitMix32 mixer
        Next b1
        b1 = 0
        b2 = 58
        b3 = 119
        b4 = 178
        initialized = True
    End If

    'Increment and generate new number
    b1 = (b1 + BIT1) Mod BIT8
    b2 = (b2 + BIT1) Mod BIT8
    b3 = (b3 + BIT1) Mod BIT8
    b4 = (b4 + BIT1) Mod BIT8
    cast = tt(b1) + tt(b2) + tt(b3) + tt(b4)
    divi = Int(cast * DIV32)
    tt(b1) = cast - divi * BIT32
    LFIB4_U32 = tt(b1)
End Function

'<========= Mersenne Twister (Mtran) code =========
Function MTran_U32^()
    'MERSENNE TWISTER pseudorandom number generator designed for older VBA limited to signed 32-bit
    'Ref. Makoto Matsumoto and Takuji Nishimura, 1996/1999
    Const kShr1& = 2            '2==2^1
    Const kShr5& = 32           '32==2^5
    Const kShr6& = 64           '64==2^6
    Const kShr11& = 2048        '2048==2^11
    Const kShr18& = 262144      '262144==2^18
    Const kShr30& = 1073741824  '1073741824==2^30  used for init.
    Const kShl7& = 128          '128==2^7
    Const kShl15& = 32768       '32768==2^15
    Dim tmp&
    Dim y&
    Dim kk&
    Dim tt&
    Dim lseed&
    Dim dseed#
    Static mt&(0 To Nuplim)     'the state vector array
    Static mti&
    Static mag01&(0 To 2)
    Static mtb As Boolean       'TRUE after initialized. (Should be =FALSE before first call)
'---
    If Not mtb Then mti = Nplus1  'mti=MTn&+1 means mt[MTn&] is not initialized
    If mti >= MTn& Then  'generate MTn& words at one time
        If mti = Nplus1 Then 'initialize
            dseed# = Now()
            lseed& = CLng((dseed# - Int(dseed#)) * 100000) '5489 is default seed in the range [-2147483648, 2147483647]
            mt(0) = lseed& And &HFFFFFFFF
            For mti = 1 To Nuplim
                tt = mt(mti - 1)
                mt(mti) = uAdd(uMult(1812433253, (tt Xor uDiv(tt, kShr30))), mti)
            Next
            mtb = True 'means mt[MTn&] is now initialized
            mag01(0) = 0
            mag01(1) = MATRIX_A
        End If
        For kk = 0 To NuplimLessM
            y = (mt(kk) And UPPER_MASK) Or (mt(kk + 1) And LOWER_MASK)
            mt(kk) = (mt(kk + MTm&) Xor uDiv(y, kShr1)) Xor mag01(y And &H1)
        Next
        For kk = kk To NuplimLess1
            y = (mt(kk) And UPPER_MASK) Or (mt(kk + 1) And LOWER_MASK)
            mt(kk) = (mt(kk + kDiffMN) Xor uDiv(y, kShr1)) Xor mag01(y And &H1)
        Next
        y = (mt(Nuplim) And UPPER_MASK) Or (mt(0) And LOWER_MASK)
        mt(Nuplim) = (mt(Muplim) Xor uDiv(y, kShr1)) Xor mag01(y And &H1)
        mti = 0
    End If
    y = mt(mti)
    mti = mti + 1
    y = (y Xor uDiv(y, kShr11))
    y = (y Xor uMult(y, kShl7) And &H9D2C5680)
    y = (y Xor uMult(y, kShl15) And &HEFC60000)
    kk& = (y Xor uDiv(y, kShr18)) 'genrand_int32SignedLong() output
    MTran_U32 = CLngLng(kk&) + CLngLng(k2_31)
End Function

Private Function uAdd&(ByVal x&, ByVal y&)
    'Helper function for MTran()
    Dim tmp#
    tmp = CDbl(x) + y
    If tmp < k2_31Neg Then
        uAdd = CLng(k2_32 + tmp)
    Else
        If tmp > k2_31b Then uAdd = CLng(tmp - k2_32) Else uAdd = CLng(tmp)
    End If
End Function 'uAdd

Private Function uMult&(ByVal x&, ByVal y&)
    'Helper function for MTran()
    Const k2_8& = 256
    Const k2_16& = 65536
    Const k2_24& = 16777216
    Dim bb&, cc&, dd&, ff&, gg&, hh&, r3&, r2&, r1&, r0&
    Dim tmp#
    bb = (x \ k2_16) Mod k2_8
    cc = (x \ k2_8) Mod k2_8
    dd = x Mod k2_8
    ff = (y \ k2_16) Mod k2_8
    gg = (y \ k2_8) Mod k2_8
    hh = y Mod k2_8
    r0 = dd * hh
    r1 = cc * hh + dd * gg + r0 \ k2_8
    r2 = bb * hh + cc * gg + dd * ff + r1 \ k2_8
    r3 = (((x \ k2_24) * hh + bb * gg + cc * ff + dd * (y \ k2_24)) Mod k2_8) + r2 \ k2_8
    tmp = CDbl(r3 Mod k2_8) * k2_24 + (r2 Mod k2_8) * k2_16 + (r1 Mod k2_8) * k2_8 + (r0 Mod k2_8)
    If tmp < k2_31Neg Then
        uMult = CLng(k2_32 + tmp)
    Else
        If tmp > k2_31b Then uMult = CLng(tmp - k2_32) Else uMult = CLng(tmp)
    End If
End Function 'uMult

Private Function uDiv&(ByVal x&, ByVal y&)
    'Helper function for MTran()
    If x < 0 Then uDiv = CLng(Fix((k2_32 + x) / y)) Else uDiv = CLng(Fix(x / y))
End Function 'uDiv
'========= end of MTran code =========>
