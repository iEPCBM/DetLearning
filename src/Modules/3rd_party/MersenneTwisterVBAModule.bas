Attribute VB_Name = "MersenneTwisterVBAModule"
Option Explicit

'******************************************************************************
' ***** COMMENTS AND REMARKS *****

' Translation by Jerry Wang as of February 2015 of the C-program MT19937
'  containing the algorithm of the Mersenne Twister Pseudorandom Number
'  Generator (MT PRNG) into an Excel/Visual Basic for Applications (Excel/VBA)
'  module "MersenneTwisterVBAModule".

' Special thanks to Professors Takuji Nishimura and Makoto Matsumoto of
'  Hiroshima University for providing the original C-program for MT19937,
'  with the latest version released as of 2006/1/26.

' One of the goals for this particular translation from C to VBA was to
'  improve performance in VBA as much as possible by minimizing the use of
'  Double Data Types to represent 32-bit Unsigned Integers, which do not exist
'  as a data type in VBA.  This was an effort to minimize the use of floating
'  point operations, which are slower and computationally more expensive than
'  bitwise and integer operations, and prone to accumulated rounding errors
'  due to the limitations of the binary number system in accurately
'  representing the decimal number system.
' The exception here are situations that involve multiplication and division
'  only by powers of two (which respectively emulate left and right bit
'  shifts).  They are unaffected by these drawbacks, and ultimately are needed
'  in this code since VBA does not have the traditional bit shift operators
'  "<<" and ">>>".

' The "MTrandom()" subroutine was written specifically to output a series of
'  random numbers from 0 to 1 inclusive to column A in Microsoft Excel.

' Translations of functions from the original C code include: "main()",
'  "init_genrand()", "init_by_array()", "genrand_int32()", "genrand_int31()",
'  "genrand_real1()", "genrand_real2()", "genrand_real3()",
'  and "genrand_res53()".
' They should function as originally intended with the exception of
'  "genrand_int32()" which was rewritten into two separate functions:
'  "genrand_signedInt32()", which returns a 32-bit Signed Long Integer but
'  otherwise exactly mirrors the original "genrand_int32()" function written
'  in C; and another function "genrand_int32()" that does return a "32-bit
'  Unsigned Long Integer" but is necessarily returned as a Double Data Type.

' Even though VBA only supports 4-byte Signed Long Data Types and not 4-byte
'  Unsigned Long Data Types, it is still possible to emulate bitwise
'  operations with a few tweaks and minimal actual arithmetic, as shown below.
'  More specifically, Signed Value Representations can be treated as Unsigned
'  representations without actually converting Signed Values into Unsigned
'  Values by only relying on Bitwise Operations (AND, NOT, OR, XOR).

' If possible, it is probably better to just rely on the MT PRNG written in C
'  or C++, which is many orders of magnitude faster than VBA even if the C/C++
'  code was not optimized.  In addition, the MT PRNG is natively implemented
'  as the default PRNG for many other software tools such as Maple, MATLAB,
'  PHP, Python, and R.  For the situations where Excel is the only option, it
'  is the sincere hope that this translation to VBA will be sufficient.

' This was written for Microsoft Excel versions 2010 and later.

' The TEMPERING_SHIFT functions were redefined as VBA functions from their
'  original #define C macros.

' In addition, the functions "adder", "rShiftBy##", and "lShiftBy##", as well
'  as the constants "Bit##" and "Bits##To##" (where ## represents various
'  numbers from 0 to 31) were not part of the original C code and are needed
'  to emulate various functionalities at various points in the VBA code.

' Future projects may involve programming the 64-bit version of the MT by
'  having VBA's Currency Data Type masquerade as a 64-bit unsigned integer.
'  In theory, this would make it possible for the 64-bit version of the MT
'  PRNG to "run" on the 32-bit version of Microsoft Excel.

' The Mersenne Twister is indeed a much higher quality pseudorandom number
'  generator whose statistical qualities are suitable for Monte Carlo
'  Simulations.  In contrast, VBA's built-in function "Rnd()" is a Linear
'  Congruential Generator that outputs a pseudorandom sequence that is not
'  acceptable for Monte Carlo simulations due to the prevalence of serial
'  correlation within the sequence.  In fact, regardless of whether or not the
'  VBA command "Randomize" was used to set the seed, "Rnd()" always outputs a
'  sequence based on the following formula: "X_n+1 =
'    (((16598013 * (16777216 * X_n)) + 12820163) mod 16777216) / 16777216".
' In addition, there are no reliable sources of information regarding Excel's
'  worksheet function "RAND()", other than the fact that it appears to be a
'  pseudorandom number generator that generates at least a 53-bit word.
'  Whether or not this is because Excel stores all numeric values as Double
'  Data Types (which has 53 bits of mantissa) or the sequence is truly derived
'  from an algorithm that works with more bits remains to be seen.  Another
'  key reason to avoid Excel's RAND() function for Monte Carlo Simulations
'  (even if it passes TestU01's BigCrush Tests and has a sufficiently long
'  period) is that there is no known way to initialize RAND's seed state (as
'  of January 2015).  Because of these aforementioned reasons, it is highly
'  recommended to avoid using Excel's "RAND()" and VBA's "Rnd()" for Monte
'  Carlo Simulations and for any other situations where high quality
'  randomness is important.

' If you have questions or concerns, please contact me at (no spaces):
'   jerry . wang . 000 @ gmail . com

' ACKNOWLEDGEMENTS:
' - Professors Takuji Nishimura and Makoto Matsumoto
' - Homepage: http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/emt.html
' - See http://en.wikipedia.org/wiki/Mersenne_twister for additional info.

' ***** End Comments And Remarks *****
'******************************************************************************'/*
' A C-program for MT19937, with initialization improved 2002/1/26.
' Coded by Takuji Nishimura and Makoto Matsumoto.
'
' Before using, initialize the state by using init_genrand(seed)
' or init_by_array(init_key, key_length).
'
' Copyright (C) 1997 - 2002, Makoto Matsumoto and Takuji Nishimura,
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions
' are met:
'
'   1. Redistributions of source code must retain the above copyright
'      notice, this list of conditions and the following disclaimer.
'
'   2. Redistributions in binary form must reproduce the above copyright
'      notice, this list of conditions and the following disclaimer in the
'      documentation and/or other materials provided with the distribution.
'
'   3. The names of its contributors may not be used to endorse or promote
'      products derived from this software without specific prior written
'      permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
' "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
' LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
' A PARTICULAR PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL THE COPYRIGHT OWNER
' OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
' EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
' PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
' PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
' LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
' NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'
' Any feedback is very welcome.
' http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/emt.html
' email: m-mat @ math.sci.hiroshima-u.ac.jp (remove space)
'*/
'******************************************************************************

' ***** BEGIN MODULE *****

' Output Pseudo Random Numbers to this Column:
Const targetOutputColumn As String = "A"
' The subroutine "main" outputs 1000 iterations each of "genrand_int32()" and
'  "genrand_real2()" to the text file "mt19937arVBAout.txt":
Const NumOutputs As Long = 1000&
Const Filename As String = "mt19937arVBAout.txt"
' Likewise, the subroutine "mainXL" outputs
'  (200 rows * 5 columns) = 1000 iterations each of "genrand_int32()" and
'  "genrand_real2()" to an Excel spreadsheet.
Const rdimension As Long = 200&
Const cdimension As Long = 5&

' MASKING BIT or BITS used in the various functions below and various remarks:
'******************************************************************************
Const Bits0To1 As Long = &H3&         ' ((2^2) - (2^0))   = 3
Const Bits0To2 As Long = &H7&         ' ((2^3) - 1)       = 7
Const Bits0To3 As Long = &HF&         ' ((2^4) - 1)       = 15
Const Bits0To4 As Long = &H1F&        ' ((2^5) - 1)       = 31
Const Bits0To6 As Long = &H7F&        ' ((2^7) - 1)       = 127
Const Bits0To8 As Long = &H1FF&       ' ((2^9) - 1)       = 511
Const Bits0To10 As Long = &H7FF&      ' ((2^11) - 1)      = 2047
Const Bits0To11 As Long = &HFFF&      ' ((2^12) - 1)      = 4095
Const Bits0To12 As Long = &H1FFF&     ' ((2^13) - 1)      = 8191
Const Bits0To13 As Long = &H3FFF&     ' ((2^14) - 1)      = 16383
Const Bits0To14 As Long = &H7FFF&     ' ((2^15) - 1)      = 32767
' Appending a '&' after the constant is required, otherwise the VBA Editor
'  treats "&HFFFF" as -1 instead of 65535.  See the link (last updated on
'  February 27, 2014) for information: http://support.microsoft.com/kb/38888
Const Bits0To15 As Long = &HFFFF&     ' ((2^16) - 1)      = 65535 (0x0000FFFF)
'******************************************************************************
Const Bits0To16 As Long = &H1FFFF     ' ((2^17) - 1)      = 131071
Const Bits0To17 As Long = &H3FFFF     ' ((2^18) - 1)      = 262143
Const Bits0To19 As Long = &HFFFFF     ' ((2^20) - 1)      = 1048575
Const Bits0To20 As Long = &H1FFFFF    ' ((2^21) - 1)      = 2097151
Const Bits0To21 As Long = &H3FFFFF    ' ((2^22) - 1)      = 4194303
Const Bits0To22 As Long = &H7FFFFF    ' ((2^23) - 1)      = 8388607
Const Bits0To23 As Long = &HFFFFFF    ' ((2^24) - 1)      = 16777215
Const Bits0To24 As Long = &H1FFFFFF   ' ((2^25) - 1)      = 33554431
Const Bits0To25 As Long = &H3FFFFFF   ' ((2^26) - 1)      = 67108863
Const Bits0To27 As Long = &HFFFFFFF   ' ((2^28) - 1)      = 268435455
Const Bits0To28 As Long = &H1FFFFFFF  ' ((2^29) - 1)      = 536870911
Const Bits0To29 As Long = &H3FFFFFFF  ' ((2^30) - 1)      = 1073741823
' This is the largest possible value that the Signed Long Data Type can take.
' "Bits0To30" and the name "LOWER_MASK" are the same as defined in the module.
Const LOWER_MASK As Long = &H7FFFFFFF ' ((2^31) - 1)      = 2147483647
'******************************************************************************
' Mask all bits except for the least significant bit and the sign bit.
Const Bits1To30 As Long = &H7FFFFFFE  ' ((2^31) - (2^1))  = 2147483646
' Mask all bits except for the 5 least significant bits and the sign bit.
Const Bits5To30 As Long = &H7FFFFFE0  ' ((2^31) - (2^5))  = 2147483616
' Mask all bits except for the 6 least significant bits and the sign bit.
Const Bits6To30 As Long = &H7FFFFFC0  ' ((2^31) - (2^6))  = 2147483584
' Mask all bits except for the 11 least significant bits and the sign bit.
Const Bits11To30 As Long = &H7FFFF800 ' ((2^31) - (2^11)) = 2147481600
' Mask all bits except for the 18 least significant bits and the sign bit.
Const Bits18To30 As Long = &H7FFC0000 ' ((2^31) - (2^18)) = 2147221504
'******************************************************************************
' This could be "Bits0To0"
Const Bit0 As Long = &H1&             ' (2^0)             = 1
Const Bit1 As Long = &H2&             ' (2^1)             = 2
Const Bit2 As Long = &H4&             ' (2^2)             = 4
Const Bit3 As Long = &H8&             ' (2^3)             = 8
Const Bit4 As Long = &H10&            ' (2^4)             = 16
Const Bit5 As Long = &H20&            ' (2^5)             = 32
Const Bit6 As Long = &H40&            ' (2^6)             = 64
Const Bit7 As Long = &H80&            ' (2^7)             = 128
Const Bit8 As Long = &H100&           ' (2^8)             = 256
Const Bit9 As Long = &H200&           ' (2^9)             = 512
Const Bit10 As Long = &H400&          ' (2^10)            = 1024
Const Bit11 As Long = &H800&          ' (2^11)            = 2048
Const Bit12 As Long = &H1000&         ' (2^12)            = 4096
Const Bit13 As Long = &H2000&         ' (2^13)            = 8192
Const Bit14 As Long = &H4000&         ' (2^14)            = 16384
' Appending a '&' after the constant is required, otherwise the VBA Editor
'  treats "&H8000" as -32768 instead of 32768.
Const Bit15 As Long = &H8000&         ' (2^15)            = 32768 (0x00008000)
'******************************************************************************
Const Bit16 As Long = &H10000         ' (2^16)            = 65536
Const Bit17 As Long = &H20000         ' (2^17)            = 131072
Const Bit18 As Long = &H40000         ' (2^18)            = 262144
Const Bit19 As Long = &H80000         ' (2^19)            = 524288
' This also represents the maximum number of rows in an Excel worksheet.
Const Bit20 As Long = &H100000        ' (2^20)            = 1048576
Const Bit21 As Long = &H200000        ' (2^21)            = 2097152
Const Bit22 As Long = &H400000        ' (2^22)            = 4194304
Const Bit23 As Long = &H800000        ' (2^23)            = 8388608
Const Bit24 As Long = &H1000000       ' (2^24)            = 16777216
Const Bit25 As Long = &H2000000       ' (2^25)            = 33554432
Const Bit26 As Long = &H4000000       ' (2^26)            = 67108864
Const Bit27 As Long = &H8000000       ' (2^27)            = 134217728
Const Bit28 As Long = &H10000000      ' (2^28)            = 268435456
Const Bit29 As Long = &H20000000      ' (2^29)            = 536870912
Const Bit30 As Long = &H40000000      ' (2^30)            = 1073741824
' "(2^31) = 2147483648" ~ This value is too large (by 1) for VBA's Long Data
'  Type.  Converting (2^31) to a 32-bit Signed Integer yields -2147483648
'  instead.  This is the smallest 2's Complement value that can be stored in a
'  Signed Long Data Type variable.
' "Bit31" and the name "UPPER_MASK" are the same as defined in the module.
' As an aside, the VBA Editor inexplicably treats "-2147483648&" as a syntax
'  error and always converts "-2147483648" to "-2147483648#" (a Double Data
'  Type value).  Thus, the setting of the constant of "UPPER_MASK" / "Bit31"
'  to "&H80000000" rather than "-2147483648#" is a way to circumvent this
'  inconvenience and guarantee the assignment of the correct data type to the
'  relevant constant.
Const UPPER_MASK As Long = &H80000000 ' This is -2147483648, not 2147483648.
'******************************************************************************

' The "Dbl" appended to the following constants signify that they are declared
'  as 8-Byte Double Data Type Constants.
' This constant has the same value as "Bit26", except that it is declared as a
'  Double Data Type in for use in the functions "genrand_res53()" and
'  "genrand_res53XL()" (Note the "#" character appended after the number).
Const Bit26Dbl As Double = 67108864#
'******************************************************************************
' The following Constants are too large to be stored in a Long Data Type and
'  must be stored in a Double Data Type even though they are Integer Valued.
'  Fortunately, they are only used in the finalizing of the outputs of the
'  newly generated random numbers as a real number [0 To 1].
' This is the smallest positive "integer" that can fit inside a Double Data
'  Type but not a Long Data Type.
Const Bit31Dbl As Double = 2147483648#            ' (2^31)        = 2147483648
Const Bit32Dbl As Double = 4294967296#            ' (2^32)        = 4294967296
' If VBA had a 4-byte Unsigned Long Integer Data Type, 4294967295 would have
'  been the largest possible integer value that could be stored in it.
Const Bit32DblLessOne As Double = Bit32Dbl# - 1#  ' (2^32) - 1    = 4294967295
'******************************************************************************
' This constant is for sake of accuracy because (in this situation at least)
'  (x / 4294967295) is NOT equal to (x * (1 / 4294967295)).  The latter has
'  been shown to produce results that are more accurate relative to the
'  results produced from the original C code.
Const Bit32DblLessOneInv As Double = 1# / Bit32DblLessOne#
'******************************************************************************
' This is the original constant: (2^53) = 9007199254740992
Const Bit53Dbl As Double = 7199254740992# + 9E+15
' Notice that:
'    "Bit53Dbl#            - 9E+15" successfully yields "7199254740992", and
'    "9.00719925474099E+15 - 9E+15" yields              "7199254740990".
' This constant is needed because while a Double Data Type variable can store
'  up to 53 significant bits mantissa, the VBA Editor cannot process more than
'  15 significant digits of any variable of any type (999,999,999,999,999 only
'  requires 50 bits).  This constant is used in the functions
'  "genrand_res53()" and "genrand_res53XL()".
'******************************************************************************

' This is the Default Seed.  4357 has also been used in the past.  This value
'  must be an integer between -2147483648 and 2147483647 inclusive.
' Note that the 10/29/1999 update to the original C code (which itself is now
'  outdated) revised the initialization routine so that an initial seed value
'  of zero is acceptable:
'    http://www.math.sci.hiroshima-u.ac.jp/
'    ~m-mat/MT/VERSIONS/C-LANG/ver991029.html
Const defaultSeed As Long = 5489&
'******************************************************************************
' #include<stdio.h>

' /* Period parameters */
' #define N 624
' #define M 397
' #define MATRIX_A 0x9908b0df   /* constant vector a */
' #define UPPER_MASK 0x80000000 /* most significant w-r bits */, defined above
' #define LOWER_MASK 0x7fffffff /* least significant r bits */, defined above
'******************************************************************************
Const constN As Long = 624&
Const constM As Long = 397&
Const constN_LessOne As Long = constN& - 1&
Const MATRIX_A As Long = &H9908B0DF

'******************************************************************************
'/* Tempering parameters */
'#define TEMPERING_MASK_B 0x9d2c5680
'#define TEMPERING_MASK_C 0xefc60000
'#define TEMPERING_SHIFT_U(y)  (y >> 11)
'#define TEMPERING_SHIFT_S(y)  (y << 7)
'#define TEMPERING_SHIFT_T(y)  (y << 15)
'#define TEMPERING_SHIFT_L(y)  (y >> 18)
'******************************************************************************
' The TEMPERING_SHIFT macro functions are defined further into the module.
Const TEMPERING_MASK_B As Long = &H9D2C5680
Const TEMPERING_MASK_C As Long = &HEFC60000
'******************************************************************************

' Global Variables (Not Quite "Static" in the traditional sense for VBA)
'******************************************************************************
'    static unsigned long mt[N]; /* the array for the state vector  */
'    static int mti=N+1; /* mti==N+1 means mt[N] is not initialized */
'******************************************************************************
' Create a length 624 array to store the state of the generator and a
'  corresponding indexing variable.
Private MT(0& To constN_LessOne&) As Long
Private mti As Long
' The following variable replaces the static functionality of the variable
'  "mti".  It is automatically initialized to "False" by default in VBA when
'  first declared globally before runtime, and will keep its value as long as
'  this Excel workbook is open.
Private mtInitialized As Boolean
'******************************************************************************

' Random Notes: (32 * 624) - 31 = 19937; (2 ^ 19937) - 1 is a Mersenne Prime.

' ***** End of Constants and Global Variables Section *****

Sub MTrandom()
'******************************************************************************
' This subroutine was written specifically for Excel.  It takes a seed and
'  generates a column of Pseudo Random Numbers.
'******************************************************************************
    ' Initial Housekeeping
    On Error GoTo Final_Housekeeping
    '**************************************************************************
    ' Array size to be determined during runtime.
    Dim arr() As Double
    Dim dimension As Long, j As Long
    '**************************************************************************
    ' Code to speed up the application.
    Dim outputColumn As String: outputColumn$ = targetOutputColumn$
    With Application
        .DisplayAlerts = False
        .EnableEvents = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    Range(outputColumn$ & ":" & outputColumn$).ClearContents
    '**************************************************************************
    ' *** MT Initialization ***
    ' Seed should be an integer between -2147483648 and 2147483647 inclusive.
    ' Otherwise, the seed will be set to the default seed...or an error will
    '  be thrown by VBA (probably either Overflow (6) or Type Mismatch (13)).
    init_genrandXL
    ' "Bit20" is 1048576; also the max number of rows in an Excel worksheet.
    dimension& = InputBox("How Many?: ", "How Many?", Bit20&)
    '**************************************************************************
    ' dimension cannot be less than 1 or greater than 1048576
    If (dimension& < 1&) Or (dimension& > Bit20&) Then
        dimension& = 1&
    End If
    ' Resize the array.  The array is required to be two-dimensional because
    '  one-dimensional arrays in VBA are "horizontal", not "vertical", and
    '  arrays in Excel are always two-dimensional.
    ReDim arr#(1& To dimension&, 1& To 1&)
    '**************************************************************************
    ' Get an array of ("dimension") Pseudo Random Numbers via MT and normalize
    '  them such that they fall between 0 and 1 inclusive.  All numerical
    '  values in an Excel spreadsheet are always stored as 8-byte Double Data
    '  Types even if they are actual integers (as of 2011, according to
    '  https://msdn.microsoft.com/en-us/library/office/
    '  bb687869%28v=office.15%29.aspx), so at this point in the code
    '  there is no real purpose in forcing a data type to a Signed Long Value.
    For j& = 1& To dimension&
        arr#(j&, 1&) = genrand_real1#
    Next
    '**************************************************************************
Final_Housekeeping:
'******************************************************************************
    ' These settings should be restored upon completion of the subroutine.
    With Application
        .CutCopyMode = False
        .DisplayAlerts = True
        .EnableEvents = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    '**************************************************************************
    ' If an error was thrown and On Error is enabled, display a message box.
    If Err.Number <> 0& Then
        MsgBox Err.Number
    Else
        '**********************************************************************
        ' Otherwise, finally, output the Newly Generated Array of MT Pseudo
        '  Random Numbers in a single operation.
        Range( _
            outputColumn$ & "1:" & outputColumn$ & CStr(dimension&)) = arr#
    End If
End Sub     'MTrandom
'******************************************************************************

Sub init_genrand(ByVal seed As Long)
'******************************************************************************
' Run this subroutine before running any of the 'genrand' functions.  If
'  running from an Excel worksheet, then run "init_genrandXL()" via the Run
'  Macro Dialog box instead (hold down "ALT" and press "F8" in Excel itself).
' Note that the obsolete 10/29/1999 update to the original C code revised the
'  initialization routine so that an initial seed value of zero is acceptable.
'******************************************************************************
' /* initializes mt[N] with a seed */
' void init_genrand(unsigned long s)
' {
'     mt[0]= s & 0xffffffffUL;
'     for (mti=1; mti<N; mti++) {
'         mt[mti] =
'         (1812433253UL * (mt[mti-1] ^ (mt[mti-1] >> 30)) + mti);
'         /* See Knuth TAOCP Vol2. 3rd Ed. P.106 for multiplier. */
'         /* In the previous versions, MSBs of the seed affect   */
'         /* only MSBs of the array mt[].                        */
'         /* 2002/01/09 modified by Makoto Matsumoto             */
'         mt[mti] &= 0xffffffffUL;
'         /* for >32 bit machines */
'     }
' }
'******************************************************************************
    Dim multiplicand As Long, addend As Long
    '**************************************************************************
    ' Unlike in the original C program, seed is an integer from -2147483648 to
    '  2147483647, instead of from 0 to 4294967295.
    MT&(0&) = seed&
    '**************************************************************************
    ' Initialize the State Array MT.
    For mti& = 1& To constN_LessOne&
        multiplicand& = MT&(mti& - 1&) Xor rShiftBy30&(MT&(mti& - 1&))
        '**********************************************************************
        ' The following lines all emulate Unsigned Multiplication of the
        '  value "multiplicand" by the number 1812433253 = 0x6C078965 =
        '  "0b 0110 1100 0000 0111 1000 1001 0110 0101 (binary)" with
        '  automatic discard of all bits except for the lowest 32 bits.
        '**********************************************************************
        ' The first line emulates: "Multiplicand Left Shifted by Zero Bits
        '  plus Multiplicand Left Shifted by Two Bits".
        addend& = adder&(multiplicand&, lShiftBy2&(multiplicand&))
        addend& = adder&(addend&, lShiftBy5&(multiplicand&))
        addend& = adder&(addend&, lShiftBy6&(multiplicand&))
        addend& = adder&(addend&, lShiftBy8&(multiplicand&))
        addend& = adder&(addend&, lShiftBy11&(multiplicand&))
        addend& = adder&(addend&, lShiftBy15&(multiplicand&))
        addend& = adder&(addend&, lShiftBy16&(multiplicand&))
        addend& = adder&(addend&, lShiftBy17&(multiplicand&))
        addend& = adder&(addend&, lShiftBy18&(multiplicand&))
        addend& = adder&(addend&, lShiftBy26&(multiplicand&))
        addend& = adder&(addend&, lShiftBy27&(multiplicand&))
        addend& = adder&(addend&, lShiftBy29&(multiplicand&))
        addend& = adder&(addend&, lShiftBy30&(multiplicand&))
        ' The following are equivalent:
        ' x * 1812433253
        ' = x * (1 + 4 + 32 + 64 + 256 + 2048 + 32768 + 65536 + 131072 +
        '    262144 + 67108864 + 134217728 + 536870912 + 1073741824)
        ' = x * ((2^0) + (2^2) + (2^5) + (2^6) + (2^8) + (2^11) + (2^15) +
        '    (2^16) + (2^17) + (2^18) + (2^26) + (2^27) + (2^29) + (2^30))
        '**********************************************************************
        'Add "mti" to the result.  "mti" should be left at 624 at looping end.
        MT&(mti&) = adder&(addend&, mti&)
    Next
    '**************************************************************************
    ' The following line replaces the functionality of this line of code that
    '  was declared in the global section of the original C code:
    '    static int mti=N+1; /* mti==N+1 means mt[N] is not initialized */
    mtInitialized = True
End Sub     'init_genrand
'******************************************************************************

Sub init_by_array(ByRef init_key() As Long)
'******************************************************************************
' Important: When calling the function "init_by_array()", the name MUST be
'  preceded by the VBA keyword "Call", and the parameter "init_key" passed to
'  it.  Otherwise, the subroutine will not function correctly.
' The parameter "key_length" no longer needs to be passed to the function.
' The array "init_key()" that is passed as a parameter MUST be declared as an
'  array of Data Type Long; otherwise, a Type Mismatch error may be thrown.
' Remember that VBA does NOT have the capability to initialize an array on a
'  single line of code (unless it was declared as Type Variant).
' The Lower Bound of the array init_key() should be set to zero.  Most of the
'  time, this should not be an issue since "Option Base 0" is the default
'  lower bound array setting in VBA.  There is some VBA code included to help
'  correct the situation if a user had chosen to declare the lower bound of
'  the init_key() array to something other than zero.
' Most importantly, remember that in C, the data types of init_key are 32-bit
'  Unsigned Long Integers, while the data types in VBA are 32-bit Signed Long
'  Integers.  Simply convert the values in the original array from their
'  Unsigned Representations [0, 4294967295] to their 2's Complement Signed
'  Representations [-2147483648, 2147483647] before passing the array as a
'  parameter to init_by_array() by subtracting 4294967296 from the Unsigned
'  Representations if their values are greater than 2147483647.
' Sign ambiguity can be mitigated somewhat by assigning hexadecimal (appended
'  by '&H') instead of decimal values to the "init_key()" array.
'******************************************************************************
' /* initialize by an array with array-length */
' /* init_key is the array for initializing keys */
' /* key_length is its length */
' /* slight change for C++, 2004/2/26 */
' void init_by_array(unsigned long init_key[], int key_length)
' {
'     int i, j, k;
'     init_genrand(19650218UL);
'     i=1; j=0;
'     k = (N>key_length ? N : key_length);2
'     for (; k; k--) {
'         mt[i] = (mt[i] ^ ((mt[i-1] ^ (mt[i-1] >> 30)) * 1664525UL))
'           + init_key[j] + j; /* non linear */
'         mt[i] &= 0xffffffffUL; /* for WORDSIZE > 32 machines */
'         i++; j++;
'         if (i>=N) { mt[0] = mt[N-1]; i=1; }
'         if (j>=key_length) j=0;
'     }
'     for (k=N-1; k; k--) {
'         mt[i] = (mt[i] ^ ((mt[i-1] ^ (mt[i-1] >> 30)) * 1566083941UL))
'           - i; /* non linear */
'         mt[i] &= 0xffffffffUL; /* for WORDSIZE > 32 machines */
'         i++;
'         if (i>=N) { mt[0] = mt[N-1]; i=1; }
'     }
'
'     mt[0] = 0x80000000UL; /* MSB is 1; assuring non-zero initial array */
'}
'******************************************************************************
    Dim i As Long, j As Long, k As Long, multiplicand As Long, _
        addend As Long, initKeyLowerBound As Long, key_length As Long
    ' Code to ensure the first element of "init_key()" is the zeroth element.
    initKeyLowerBound& = LBound(init_key&)
    ' Passing "key_length" as in the original C code is unnecessary since the
    '  VBA functions "UBound()" and "LBound()" are built-in.
    key_length& = UBound(init_key&) - initKeyLowerBound& + 1&
    init_genrand 19650218
    i& = 1&
    j& = 0&
    If constN& > key_length& Then
        k& = constN&
    Else
        k& = key_length&
    End If
    '**************************************************************************
    For k& = k& To 1& Step -1&
        multiplicand& = MT&(i& - 1&) Xor rShiftBy30&(MT&(i& - 1&))
        '**********************************************************************
        ' The following lines emulate multiplying the multiplicand by 1664525
        '  with built-in 32-bit overflow protection:
        addend& = adder&(multiplicand&, lShiftBy2&(multiplicand&))
        addend& = adder&(addend&, lShiftBy3&(multiplicand&))
        addend& = adder&(addend&, lShiftBy9&(multiplicand&))
        addend& = adder&(addend&, lShiftBy10&(multiplicand&))
        addend& = adder&(addend&, lShiftBy13&(multiplicand&))
        addend& = adder&(addend&, lShiftBy14&(multiplicand&))
        addend& = adder&(addend&, lShiftBy16&(multiplicand&))
        addend& = adder&(addend&, lShiftBy19&(multiplicand&))
        addend& = adder&(addend&, lShiftBy20&(multiplicand&))
        '**********************************************************************
        ' Add both init_key(j) and j to the final result:
        MT&(i&) = adder&(adder&(MT&(i&) Xor addend&, _
            init_key&(j& + initKeyLowerBound&)), j&)
        i& = i& + 1&
        j& = j& + 1&
        If i& >= constN& Then
            MT&(0&) = MT&(constN_LessOne&)
            i& = 1&
        End If
        If j& >= key_length& Then
            j& = 0&
        End If
    Next
    '**************************************************************************
    For k& = constN_LessOne& To 1& Step -1&
        multiplicand& = MT&(i& - 1&) Xor rShiftBy30&(MT&(i& - 1&))
        '**********************************************************************
        ' The following lines emulate multiplying the multiplicand by
        '  1566083941 with built-in 32-bit overflow protection:
        addend& = adder&(multiplicand&, lShiftBy2&(multiplicand&))
        addend& = adder&(addend&, lShiftBy5&(multiplicand&))
        addend& = adder&(addend&, lShiftBy6&(multiplicand&))
        addend& = adder&(addend&, lShiftBy8&(multiplicand&))
        addend& = adder&(addend&, lShiftBy9&(multiplicand&))
        addend& = adder&(addend&, lShiftBy11&(multiplicand&))
        addend& = adder&(addend&, lShiftBy15&(multiplicand&))
        addend& = adder&(addend&, lShiftBy19&(multiplicand&))
        addend& = adder&(addend&, lShiftBy20&(multiplicand&))
        addend& = adder&(addend&, lShiftBy22&(multiplicand&))
        addend& = adder&(addend&, lShiftBy24&(multiplicand&))
        addend& = adder&(addend&, lShiftBy26&(multiplicand&))
        addend& = adder&(addend&, lShiftBy27&(multiplicand&))
        addend& = adder&(addend&, lShiftBy28&(multiplicand&))
        addend& = adder&(addend&, lShiftBy30&(multiplicand&))
        '**********************************************************************
        ' Subtract "i" from the final result:
        MT&(i&) = adder&(MT&(i&) Xor addend&, (-1&) * i&)
        i& = i& + 1&
        If i& >= constN& Then
            MT&(0&) = MT&(constN_LessOne&)
            i& = 1&
        End If
    Next
    '**************************************************************************
    ' A non-zero array is assured.
    MT&(0&) = UPPER_MASK&
End Sub     'init_by_array
'******************************************************************************

Function genrand_signedInt32() As Long
'******************************************************************************
' This is the functional translation of the original C-code function of
'  "genrand_int32()".
' The return data type here is a 32-bit Signed Long value from -2147483648 to
'  2147483647 as opposed to a 32-bit Unsigned Long value from 0 to 4294967295.
' Semantically speaking, the function still generates a random integer on the
'  [0,0xffffffff]-interval (0xffffffff represents -1 instead of 4294967295).
' The basic idea remains unchanged: "genrand_signedInt32()" will return a
'  32-bit Integer that is Pseudorandom-Uniformly Distributed from a Minimum
'  Bound to a Maximum Bound.  Both functions have 4294967296 possible output
'  values, all of which occur with pseudo-equal likelihood.
' In the original code, there was a 2-by-1 static array declared in
'  "genrand()" as mag01, with mag01(0) set to 0 and mag01(1) set to the value
'  of "MATRIX_A".  This was recoded as shown below.
' Note that declaring an unsigned long variable "y" is unnecessary because the
'  return value and function share the same name of "genrand_signedInt32".
'******************************************************************************
' /* generates a random number on [0,0xffffffff]-interval */
' unsigned long genrand_int32(void)
' {
'     unsigned long y;
'     static unsigned long mag01[2]={0x0UL, MATRIX_A};
'     /* mag01[x] = x * MATRIX_A  for x=0,1 */
'
'     if (mti >= N) { /* generate N words at one time */
'         int kk;
'
'         if (mti == N+1)   /* if init_genrand() has not been called, */
'             init_genrand(5489UL); /* a default initial seed is used */
'
'         for (kk=0;kk<N-M;kk++) {
'             y = (mt[kk]&UPPER_MASK)|(mt[kk+1]&LOWER_MASK);
'             mt[kk] = mt[kk+M] ^ (y >> 1) ^ mag01[y & 0x1UL];
'         }
'         for (;kk<N-1;kk++) {
'             y = (mt[kk]&UPPER_MASK)|(mt[kk+1]&LOWER_MASK);
'             mt[kk] = mt[kk+(M-N)] ^ (y >> 1) ^ mag01[y & 0x1UL];
'         }
'         y = (mt[N-1]&UPPER_MASK)|(mt[0]&LOWER_MASK);
'         mt[N-1] = mt[M-1] ^ (y >> 1) ^ mag01[y & 0x1UL];
'
'         mti = 0;
'     }
'
'     y = mt[mti++];
'
'     /* Tempering */
'     y ^= (y >> 11);
'     y ^= (y << 7) & 0x9d2c5680UL;
'     y ^= (y << 15) & 0xefc60000UL;
'     y ^= (y >> 18);
'
'     return y;
'}
'******************************************************************************
    ' The following If Statement replaces the approach in the original C code
    '  of initializing the static variable "mti" to be equal to "N+1" as a
    '  check for the initialization of the MT array, since VBA does not have
    '  an equivalent "static" keyword to C:
    If Not mtInitialized Then
        init_genrand defaultSeed&
    End If
    '**************************************************************************
    '{ /* generate N words at one time */
    ' This is executed every "N" numbers.  Since "mti" is no longer
    '  initialized to "N+1", it is not necessary to check if "mti" is greater
    '  than "constN".
    If mti& = constN& Then
        '**********************************************************************
        ' The following lines were replaced earlier by a different approach
        '  that occurs outside of the current If Statement since VBA does not
        '  have an equivalent "static" keyword to C:
        '    if (mti == N+1)   /* if init_genrand() has not been called, */
        '        init_genrand(5489UL); /* a default initial seed is used */
        ' This also means that declaring the variable "kk" is unnecessary.
        '**********************************************************************
        ' Alternate Pseudocode section taken from
        '  http://en.wikipedia.org/wiki/Mersenne_twister, June 2014:
        ' // Generate an array of 624 untempered numbers
        ' function generate_numbers() {
        '     for i from 0 to 623 {
        '         // bit 31 (32nd bit) of MT[i]
        '         int y := (MT[i] and 0x80000000)
        '             // bits 0-30 (first 31 bits) of MT[...]
        '             or (MT[(i+1) mod 624] and 0x7fffffff)
        '         MT[i] := MT[(i + 397) mod 624] xor (right shift by 1 bit(y))
        '         if (y mod 2) != 0 { // y is odd
        '             MT[i] := MT[i] xor (2567483615) // 0x9908b0df
        '         }
        '     }
        ' }
        '**********************************************************************
        For mti& = 0& To constN_LessOne&
            genrand_signedInt32& = (MT&(mti&) And UPPER_MASK&) Or _
                (MT&((mti& + 1&) Mod constN&) And LOWER_MASK&)
            '******************************************************************
            MT&(mti&) = MT&((mti& + constM&) Mod constN&) Xor _
                rShiftBy1&(genrand_signedInt32&)
            '******************************************************************
            ' If y is Odd; emulates "mag01[y & 0x1]" in the original C code.
            If CBool(genrand_signedInt32& And Bit0&) Then
                ' MATRIX_A is originally from the static unsigned long array
                '  mag01[2] = {0x0, MATRIX_A} in the original C code.
                MT&(mti&) = MT&(mti&) Xor MATRIX_A&
            End If
            '******************************************************************
        Next
        ' Emulates "mti = 0;"
        mti& = 0&
        '**********************************************************************
    End If
    
    '**************************************************************************
    ' Extract a Tempered Pseudorandom Number based on the index-th value.
    ' "genrand_signedInt32" is used in place of the variable "y" here.
    genrand_signedInt32& = MT&(mti&)
    ' Emulates 'mti++', since it cannot be combined with the previous line on
    '  a single line of code (C/C++ style) in VB/VBA.
    mti& = mti& + 1&
    '**************************************************************************
    ' Carry out Tempering Shifts and Tempering Masks:
    genrand_signedInt32& = genrand_signedInt32& Xor _
        rShiftBy11&(genrand_signedInt32&)
    genrand_signedInt32& = genrand_signedInt32& Xor _
        (lShiftBy7&(genrand_signedInt32&) And TEMPERING_MASK_B&)
    genrand_signedInt32& = genrand_signedInt32& Xor _
        (lShiftBy15&(genrand_signedInt32&) And TEMPERING_MASK_C&)
    genrand_signedInt32& = genrand_signedInt32& Xor _
        rShiftBy18&(genrand_signedInt32&)
    '**************************************************************************
    
    ' A Return statement "return y;" is not needed because the last value of
    '  "genrand_signedInt32" will be returned automatically.
    ' Tbe returned value of "genrand_signedInt32()" will be between
    '  -2147483648 and 2147483647 inclusive.
    ' The name change from "genrand_int32" to "genrand_signedInt32" EMPHASIZES
    '  the fact that the function returns a SIGNED Long value.
End Function    'genrand_signedInt32
'******************************************************************************

Function genrand_int32() As Double
'******************************************************************************
' This function converts the 2's Complement Signed 32-bit Integer output
'  from the "genrand_signedInt32()" function into its Unsigned Value
'  representation which will take on values from 0 to 4294967295.  The return
'  data type of this function is necessarily Double since VBA does not have a
'  32-bit Unsigned Long Integer data type.  This function exactly replicates
'  the output values (but not the functionality) of "genrand_int32(void)" from
'  the original C code, except it returns Double Data Type variables.
'******************************************************************************
' /* generates a random number on [0,0xffffffff]-interval */
' unsigned long genrand_int32(void)
'******************************************************************************
    genrand_int32# = CDbl(genrand_signedInt32&)
    If genrand_int32# < 0# Then
        genrand_int32# = genrand_int32# + Bit32Dbl#
    End If
End Function    'genrand_int32
'******************************************************************************

Function genrand_int31() As Long
'******************************************************************************
' This function returns a Pseudo Random Unsigned 31-bit Integer by logically
'  right-shifting a Pseudo Random 32-bit Signed Integer Value by one bit.
'  Since the final value is an integer between 0 and 2147483647 inclusive, the
'  return data type here does not need to be at Double precision.
'******************************************************************************
' /* generates a random number on [0,0x7fffffff]-interval */
' long genrand_int31(void)
' {
'     return (long)(genrand_int32()>>1);
' }
'******************************************************************************
    ' Emulation of logical right shifts automatically sets the return value's
    '  sign bit to zero; thus, there are no problems using the function
    '  "genrand_signedInt32()" in place of "genrand_Int32()" in order to
    '  maximize the performance of the function.
    genrand_int31& = rShiftBy1&(genrand_signedInt32&)
End Function    'genrand_int31
'******************************************************************************

Function genrand_real1() As Double
'******************************************************************************
' This function converts the 2's Complement Signed 32-bit Integer output
'  from the "genrand_signedInt32()" function into its Unsigned Value
'  representation and normalizes the values from integers [0, 4294967295] to
'  the real numbers zero inclusive to one inclusive.
' Note that because "genrand_signedInt32()" has an even number of possible
'  output values, and because of the multiplication by (1.0 / 4294967295.0)
'  that makes the output value of exactly 1.0 possible, this function will
'  never return a value of exactly 0.5.  On the other hand, this function has
'  exactly 2147483648 possible distinct output values that are greater than
'  0.5 (including 1.0), and exactly 2147483648 possible distinct output values
'  that are less than 0.5 (including 0.0), all occurring with pseudo-equal
'  likelihood.
'******************************************************************************
' /* generates a random number on [0,1]-real-interval */
' double genrand_real1(void)
' {
'     return genrand_int32()*(1.0/4294967295.0);
'     /* divided by 2^32-1 */
' }
'******************************************************************************
    genrand_real1# = CDbl(genrand_signedInt32&)
    ' Multiplication by "Bit32DblLessOneInv" is shown to be more accurate
    '  relative to the outputs from the original C code instead of Division
    '  by "Bit32DblLessOne".
    If genrand_real1# < 0# Then
        genrand_real1# = (genrand_real1# + Bit32Dbl#) * Bit32DblLessOneInv#
    Else
        genrand_real1# = genrand_real1# * Bit32DblLessOneInv#
    End If
End Function    'genrand_real1
'******************************************************************************

Function genrand_real2() As Double
'******************************************************************************
' This function converts the 2's Complement Signed 32-bit Integer output
'  from the "genrand_signedInt32()" function into its Unsigned Value
'  representation and normalizes the values from integers [0, 4294967295] to
'  the real numbers zero inclusive to one exclusive (or 0 to
'  (4294967295 / 4294967296) = 0.99999999976716935634613037109375).
'******************************************************************************
' /* generates a random number on [0,1)-real-interval */
' double genrand_real2(void)
' {
'     return genrand_int32()*(1.0/4294967296.0);
'     /* divided by 2^32 */
' }
'******************************************************************************
    genrand_real2# = CDbl(genrand_signedInt32&)
    If genrand_real2# < 0# Then
        genrand_real2# = (genrand_real2# + Bit32Dbl#) / Bit32Dbl#
    Else
        genrand_real2# = genrand_real2# / Bit32Dbl#
    End If
End Function    'genrand_real2
'******************************************************************************

Function genrand_real3() As Double
'******************************************************************************
' This function converts the 2's Complement Signed 32-bit Integer output
'  from the "genrand_signedInt32()" function into its Unsigned Value
'  representation and normalizes the values from integers [0, 4294967295] to
'  the real numbers zero exclusive to one exclusive (or
'  (0.5 / 4294967296) =          0.000000000116415321826934814453125 to
'  (4294967295.5 / 4294967296) = 0.999999999883584678173065185546875).
'******************************************************************************
' /* generates a random number on (0,1)-real-interval */
' double genrand_real3(void)
' {
'     return (((double)genrand_int32()) + 0.5)*(1.0/4294967296.0);
'     /* divided by 2^32 */
' }
'******************************************************************************
    genrand_real3# = CDbl(genrand_signedInt32&)
    If genrand_real3# < 0# Then
        genrand_real3# = ((genrand_real3# + Bit32Dbl#) + 0.5) / Bit32Dbl#
    Else
        genrand_real3# = (genrand_real3# + 0.5) / Bit32Dbl#
    End If
End Function    'genrand_real3
'******************************************************************************

Function genrand_res53() As Double
'******************************************************************************
' This function converts the 2's Complement Signed 32-bit Integer output
'  from the "genrand_signedInt32()" function into its Unsigned Value
'  representation and normalizes the values from integers [0, 4294967295] to
'  the real numbers zero inclusive to one exclusive with 53-bit resolution (or
'  0 to (9007199254740991 / 9007199254740992) =
'  0.99999999999999988897769753748434595763683319091796875).
' The return data type is necessarily Double precision, and there are exactly
'  (2^53) = 9007199254740992 possible output values.
'******************************************************************************
' /* generates a random number on [0,1) with 53-bit resolution*/
' double genrand_res53(void)
' {
'     unsigned long a=genrand_int32()>>5, b=genrand_int32()>>6;
'     return(a*67108864.0+b)*(1.0/9007199254740992.0);
' }
' /* These real versions are due to Isaku Wada, 2002/01/09 added */
'******************************************************************************
    ' Applying logical right shifts to the outputs of the function
    '  "genrand_signedInt32()" automatically discards their sign bits, so
    '  there is no need for additional code to ensure that all relevant values
    '  are Unsigned.  Furthermore, because the sign bit is discarded, the
    '  following lines of code will return identical results regardless of
    '  whether "genrand_signedInt32()" or "genrand_int32()" is used.  There
    '  should be no issues using the former for speed purposes here.
    genrand_res53# = CDbl(rShiftBy5&(genrand_signedInt32&)) * Bit26Dbl#
    genrand_res53# = (CDbl(rShiftBy6&(genrand_signedInt32&)) + _
        genrand_res53#) / Bit53Dbl#
    ' (32 - 5 + 26) = 53; the second call of "genrand_signedInt32()" resolves
    '  the zeroes left shifted in by the first call of "genrand_signedInt32()"
    '  and the subsequent bit shift operations.
    ' The User-Defined Function "adder" is not used here because it only
    '  processes Long Data types and the return value of this function is
    '  ultimately a Double Data Type.  Overflow protection is also not needed
    '  here since this function processes exactly 53 bits of mantissa.
End Function    'genrand_res53
'******************************************************************************

Sub main()
'******************************************************************************
' This subroutine more or less replicates the original "Main" function.
'******************************************************************************
' int main(void)
' {
'     int i;
'     unsigned long init[4]={0x123, 0x234, 0x345, 0x456}, length=4;
'     init_by_array(init, length);
'     printf("1000 outputs of genrand_int32()\n");
'     for (i=0; i<1000; i++) {
'       printf("%10lu ", genrand_int32());
'       if (i%5==4) printf("\n");
'     }
'     printf("\n1000 outputs of genrand_real2()\n");
'     for (i=0; i<1000; i++) {
'       printf("%10.8f ", genrand_real2());
'       if (i%5==4) printf("\n");
'     }
'     return 0;
' }
'******************************************************************************
' Initial Housekeeping
    'On Error GoTo Final_Housekeeping
    '**************************************************************************
    Dim x As Double
    Dim i As Long
    '**************************************************************************
    ' Code to speed up the application.
    With Application
        .DisplayAlerts = False
        .EnableEvents = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    '**************************************************************************
    ' Open a text file for output.  An alternate option is to use
    '  "Debug.Print"; however, this will limit the output to fewer lines of
    '  text than the subroutine will display in the Immediate Window.
    ' Occasionally, you may run into a "File already open" error.  This can be
    '  fixed by exiting and reopening Excel.
    Open Application.ThisWorkbook.Path & "\" & Filename$ For Output As #1
    ' MT Initialization
    ' "init_arr()" is declared with bounds zero to three by necessity.  Since
    '  "Option Base 0" is enabled by default in VBA, declaring the "init_arr"
    '  array with just the parameter "3&" is sufficient (as opposed to
    '  "0& To 3&").  There is additional code to automatically account for the
    '  user that decides to use a lower bound other than zero.
    ' VBA does not have the capability to initialize an array of any data type
    '  other than Variant on a single line.  Feel free to modify at one's
    '  discretion to enable this functionality if one prefers this route.
    Dim init_arr(1& To 4&) As Long
    init_arr&(1&) = &H123&
    init_arr&(2&) = &H234&
    init_arr&(3&) = &H345&
    init_arr&(4&) = &H456&
    ' "key_length" does not need to be passed as a parameter to
    '  "init_by_array()" as in the original code since VBA has built-in
    '  functions to retrieve the dimensions of an array.
    Call init_by_array(init_arr&)
    '**************************************************************************
    ' Generate Numbers
    ' The results of "genrand_int32()" and "genrand_real2()" needs to be
    '  stored to a local variable before applying formatting in order to
    '  display more than seven digits of significance.
    Print #1, CStr(NumOutputs&) & " outputs of genrand_int32()"
    For i& = 0& To (NumOutputs& - 1&)
        ' The output values are Unsigned 32-bit Pseudorandom "Integers".
        x# = genrand_int32#
        Print #1, Format$(x#, "##########");
        ' Align the output values
        If Len(CStr(x#)) < 10& Then
            Print #1, Space$(10& - Len(CStr(x#)));
        End If
        If (i& Mod 5&) = 4& Then
            Print #1, " "
        Else
            ' The semi-colon tells "Print" not to append a newline character.
            Print #1, " ";
        End If
    Next
    Print #1, vbNewLine & CStr(NumOutputs&) & " outputs of genrand_real2()"
    ' The numbers in the original C output text file were rounded to 8 digits.
    For i& = 0& To (NumOutputs& - 1&)
        ' The output values are random numbers on the [0, 1) interval.
        x# = genrand_real2#
        Print #1, Format$(x#, "0.00000000");
        If (i& Mod 5&) = 4& Then
            Print #1, " "
        Else
            ' The semi-colon tells "Print" not to append a newline character.
            Print #1, " ";
        End If
    Next
    Close #1
Final_Housekeeping:
'******************************************************************************
    ' These settings should be restored upon completion of the subroutine.
    With Application
        .CutCopyMode = False
        .DisplayAlerts = True
        .EnableEvents = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    '**************************************************************************
    ' If an error was thrown and On Error is enabled, display a message box.
    If Err.Number <> 0& Then
        MsgBox Err.Number
    End If
End Sub     'Main
'******************************************************************************

Sub init_genrandXL()
'******************************************************************************
' This subroutine serves as an addition to allow Excel to seed the MT PRNG
'  from outside the module by running it from Excel's View Macros Dialog Box.
'  This can be accessed by holding the "ALT" key and pressing the "F8" key,
'  typing in "init_genrandXL" into the dialog box (as opposed to just
'  "init_genrand", which will not work), and clicking "Run".  It simply
'  invokes a dialog box that accepts a numeric input and passes it as a
'  parameter to the original "init_genrand()" function.  In addition, it also
'  conveniently uses Excel's function "RAND()' to generate the default input
'  seed value each time the subroutine is run (NOT VBA's "Rnd()", which is
'  only based on a 24-bit word while "init_genrand()" is capable of accepting
'  32-bit integers as initial seed values.  As always, acceptable input values
'  are [-2147483648, 2147483647].
'******************************************************************************
    init_genrand CLng(InputBox("Input Seed: ", "Input Seed", _
        CLng(Round((Application.Evaluate("=RAND()") * _
        Bit32DblLessOne#) - Bit31Dbl#, 0&))))
End Sub     'init_genrandXL
'******************************************************************************

' The following functions convert the outputs from the original
'  "genrand_signedInt32()" function to various number ranges such as a real
'  number from 0 to 1.  Note that some of them have the addition of the
'  "Application.Volatile (True)" Flag to force them to recalculate every time
'  the Excel workbook is recalculated.  This feature necessarily comes with
'  noticeable degradations in computational performance and thus should only
'  be used within Excel formulae and not internally within VBA modules.
'  Remember to call "init_genrand()" or "init_genrandXL()" to seed the
'  Mersenne Twister before calling any of the following functions:
'******************************************************************************

Function genrand_signedInt32XL() As Double
'******************************************************************************
' This function is the same as "genrand_signedInt32()" with the addition of
'  the "Application.Volatile (True)" Flag.
' Even though "genrand_signedInt32()" returns Long Data Type outputs with no
'  issues, this function is coded to additionally and explicitly convert the
'  results of "genrand_signedInt32()" automatically to Double Data Types since
'  all numerical values in an Excel spreadsheet are always stored as Double
'  Data Types, even if they are Integers.
'******************************************************************************
    Application.Volatile (True)
    genrand_signedInt32XL# = CDbl(genrand_signedInt32&)
End Function    'genrand_signedInt32XL
'******************************************************************************

Function genrand_int32XL() As Double
'******************************************************************************
' This function is the same as "genrand_int32()" with the addition of the
'  "Application.Volatile (True)" Flag.
'******************************************************************************
    Application.Volatile (True)
    genrand_int32XL# = CDbl(genrand_signedInt32&)
    If genrand_int32XL# < 0# Then
        genrand_int32XL# = genrand_int32XL# + Bit32Dbl#
    End If
End Function    'genrand_int32XL
'******************************************************************************

Function genrand_int31XL() As Double
'******************************************************************************
' This function is the same as "genrand_int31()" with the addition of the
'  "Application.Volatile (True)" Flag.
' While "genrand_int31()" and "genrand_int31XL()" are both able to return a
'  Long Data Type outputs with no issues, the return Data Type of
'  "genrand_int31XL()" has been declared as Double since numerical values in
'  Excel spreadsheets are always stored as Double Data Types.  Thus, the
'  built-in function "CDbl()" applies here.
'******************************************************************************
    Application.Volatile (True)
    genrand_int31XL# = CDbl(rShiftBy1&(genrand_signedInt32&))
End Function
'******************************************************************************

Function genrand_real1XL() As Double
'******************************************************************************
' This function is the same as "genrand_real1()" with the addition of the
'  "Application.Volatile (True)" Flag.
'******************************************************************************
    Application.Volatile (True)
    genrand_real1XL# = CDbl(genrand_signedInt32&)
    If genrand_real1XL# < 0# Then
        genrand_real1XL# = _
            (genrand_real1XL# + Bit32Dbl#) * Bit32DblLessOneInv#
    Else
        genrand_real1XL# = genrand_real1XL# * Bit32DblLessOneInv#
    End If
End Function    'genrand_real1XL
'******************************************************************************

Function genrand_real2XL() As Double
'******************************************************************************
' This function is the same as "genrand_real2()" with the addition of the
'  "Application.Volatile (True)" Flag.
'******************************************************************************
    Application.Volatile (True)
    genrand_real2XL# = CDbl(genrand_signedInt32&)
    If genrand_real2XL# < 0# Then
        genrand_real2XL# = (genrand_real2XL# + Bit32Dbl#) / Bit32Dbl#
    Else
        genrand_real2XL# = genrand_real2XL# / Bit32Dbl#
    End If
End Function    'genrand_real2XL
'******************************************************************************

Function genrand_real3XL() As Double
'******************************************************************************
' This function is the same as "genrand_real3()" with the addition of the
'  "Application.Volatile (True)" Flag.
'******************************************************************************
    Application.Volatile (True)
    genrand_real3XL# = CDbl(genrand_signedInt32&)
    If genrand_real3XL# < 0# Then
        genrand_real3XL# = ((genrand_real3XL# + Bit32Dbl#) + 0.5) / Bit32Dbl#
    Else
        genrand_real3XL# = (genrand_real3XL# + 0.5) / Bit32Dbl#
    End If
End Function    'genrand_real3XL
'******************************************************************************

Function genrand_res53XL() As Double
'******************************************************************************
' This function is the same as "genrand_res53()" with the addition of the
'  "Application.Volatile (True)" Flag.
'******************************************************************************
    Application.Volatile (True)
    genrand_res53XL# = CDbl(rShiftBy5&(genrand_signedInt32&)) * Bit26Dbl#
    genrand_res53XL# = (CDbl(rShiftBy6&(genrand_signedInt32&)) + _
        genrand_res53XL#) / Bit53Dbl#
End Function    'genrand_res53XL
'******************************************************************************

Sub mainXL()
'******************************************************************************
' This subroutine more or less replicates the original Main function, writing
'  the outputs to the Excel spreadsheet instead of a text file.
'******************************************************************************
' Initial Housekeeping
    On Error GoTo Final_Housekeeping
    '**************************************************************************
    Dim arr1(1& To rdimension&, 1& To cdimension&) As Double
    Dim arr2(1& To rdimension&, 1& To cdimension&) As Double
    Dim i As Long, j As Long
    '**************************************************************************
    ' Code to speed up the application.
    With Application
        .DisplayAlerts = False
        .EnableEvents = False
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    Worksheets("Compare To Original Output").Range("$A:$E").ClearContents
    '**************************************************************************
    ' MT Initialization
    ' VBA does not have the capability to initialize an array of any data type
    '  other than Variant on a single line.  Feel free to modify at one's
    '  discretion to enable this functionality if one prefers this route.
    Dim init_arr(3&) As Long
    init_arr&(0&) = &H123&
    init_arr&(1&) = &H234&
    init_arr&(2&) = &H345&
    init_arr&(3&) = &H456&
    Call init_by_array(init_arr&)
    '**************************************************************************
    ' Generate Numbers
    ' The numbers in the original C output text file were rounded to 8 digits.
    ' Everything is explicitly outputted to double data types since the output
    '  medium is an Excel spreadsheet itself.
    For i& = 1& To rdimension&
        For j& = 1& To cdimension&
            ' The output values are Unsigned 32-bit Pseudorandom "Integers".
            arr1#(i&, j&) = genrand_int32#
        Next
    Next
    For i& = 1& To rdimension&
        For j& = 1& To cdimension&
            ' The output values are random numbers on the [0, 1) interval.
            arr2#(i&, j&) = Round(genrand_real2#, 8&)
        Next
    Next
Final_Housekeeping:
'******************************************************************************
    ' These settings should be restored upon completion of the subroutine.
    With Application
        .CutCopyMode = False
        .DisplayAlerts = True
        .EnableEvents = True
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    '**************************************************************************
    ' If an error was thrown and On Error is enabled, display a message box.
    If Err.Number <> 0& Then
        MsgBox Err.Number
    Else
        '**********************************************************************
        ' Otherwise, finally, output the Newly Generated Arrays of MT Pseudo
        '  Random Numbers.
        Range("$A$1") = "1000 outputs of genrand_int32()"
        Range("$A$2:$E$" & CStr(rdimension& + 1&)) = arr1#
        Range("$A$" & CStr(rdimension& + 3&)) = _
            "1000 outputs of genrand_real2()"
        Range("$A$" & CStr(rdimension& + 4&) & ":$E$" & _
            CStr((2& * (rdimension& + 1&)) + 1&)) = arr2#
    End If
End Sub     'MainXL
'******************************************************************************

Function getCoinFlip() As Boolean
'******************************************************************************
' This function returns "False" if "genrand_signedInt32()" returns
'  [-2147483648, -1] and "True" if it returns [0, 2147483647] (both outcomes
'  have 2147483648 possible intermediate values).
'******************************************************************************
    If genrand_signedInt32& < 0& Then
        getCoinFlip = False
    Else
        getCoinFlip = True
    End If
End Function    'getCoinFlip
'******************************************************************************

Function getCoinFlipXL() As Boolean
'******************************************************************************
' This function is the same as "getCoinFlip()" with the addition of the
'  "Application.Volatile (True)" Flag.
'******************************************************************************
    Application.Volatile (True)
    If genrand_signedInt32& < 0& Then
        getCoinFlipXL = False
    Else
        getCoinFlipXL = True
    End If
End Function    'getCoinFlipXL
'******************************************************************************

Function getRndNormalStandard() As Double
'******************************************************************************
' This function uses the result of "genrand_signedInt32()" normalized to
'  [0, 1] in conjunction with Excel's Inverse Normal Distribution function to
'  generate a Normally Distributed Pseudo Random Number with Zero Mean and
'  Unit Standard Deviation.  Rigorous speed tests have not been carried out on
'  the use of Excel's Inverse Normal Distribution function for this purpose.
' In general, there is little incentive to create a User-Defined Function to
'  replace an already built-in, native Excel function, since native built-in
'  functions are almost always faster and more reliable than User-Defined
'  Functions with the same functionality.  This is different from the
'  situation of creating the Mersenne Twister for Excel/VBA to replace the
'  Worksheet Function "=RAND()" and the Excel VBA Function "Rnd()", since the
'  former has been shown to be far more reliable than the latter.
' Note that because "genrand_signedInt32()" has an even number of possible
'  output values, and because of the multiplication by (1.0 / 4294967295.0)
'  that makes the intermediate output value of exactly 1.0 possible, this
'  function is unable to return a value that is exactly equal to the value of
'  0.0.  On the other hand, there are exactly 2147483648 possible distinct
'  output values that are greater than 0.0, and exactly 2147483648 possible
'  distinct output values that are less than 0.0.
'******************************************************************************
    getRndNormalStandard# = CDbl(genrand_signedInt32&)
    If getRndNormalStandard# < 0# Then
        getRndNormalStandard# = WorksheetFunction.Norm_Inv#( _
            (getRndNormalStandard# + Bit32Dbl#) * Bit32DblLessOneInv#, 0#, 1#)
    Else
        getRndNormalStandard# = WorksheetFunction.Norm_Inv#( _
            getRndNormalStandard# * Bit32DblLessOneInv#, 0#, 1#)
    End If
End Function    'getRndNormalStandard
'******************************************************************************

Function getRndNormalStandardXL() As Double
'******************************************************************************
' This function is the same as "getRndNormalStandard()" with the addition of
'  the "Application.Volatile (True)" Flag.
'******************************************************************************
    Application.Volatile (True)
    getRndNormalStandardXL# = CDbl(genrand_signedInt32&)
    If getRndNormalStandardXL# < 0# Then
        getRndNormalStandardXL# = WorksheetFunction.Norm_Inv#( _
            (getRndNormalStandardXL# + Bit32Dbl#) * _
            Bit32DblLessOneInv#, 0#, 1#)
    Else
        getRndNormalStandardXL# = WorksheetFunction.Norm_Inv#( _
            getRndNormalStandardXL# * Bit32DblLessOneInv#, 0#, 1#)
    End If
End Function    'getRndNormalStandardXL
'******************************************************************************

Function getRndNormalDist(Optional ByVal mean As Double = 0#, _
    Optional ByVal std As Double = 1#) As Double
'******************************************************************************
' This function uses the result of "genrand_signedInt32()" normalized to
'  [0, 1] in conjunction with Excel's Inverse Normal Distribution function to
'  generate a Normally Distributed Pseudo Random Number with the Mean and
'  Standard Deviations passed as optional parameters.
' Note that because "genrand_signedInt32()" has an even number of possible
'  output values, and because of the multiplication by (1.0 / 4294967295.0)
'  that makes the intermediate output value of exactly 1.0 possible, this
'  function is unable to return a value that is exactly equal to the value of
'  "mean".  On the other hand, there are exactly 2147483648 possible distinct
'  output values that are greater than "mean", and exactly 2147483648 possible
'  distinct output values that are less than "mean".
'******************************************************************************
    getRndNormalDist# = CDbl(genrand_signedInt32&)
    If getRndNormalDist# < 0# Then
        getRndNormalDist# = WorksheetFunction.Norm_Inv#((getRndNormalDist# _
            + Bit32Dbl#) * Bit32DblLessOneInv#, mean#, Abs(std#))
    Else
        getRndNormalDist# = WorksheetFunction.Norm_Inv#( _
            getRndNormalDist# * Bit32DblLessOneInv#, mean#, Abs(std#))
    End If
End Function    'getRndNormalDist
'******************************************************************************

Function getRndNormalDistXL(Optional ByVal mean As Double = 0#, _
    Optional ByVal std As Double = 1#) As Double
'******************************************************************************
' This function is the same as "getRndNormalDist()" with the addition of the
'  "Application.Volatile (True)" Flag.
'******************************************************************************
    Application.Volatile (True)
    getRndNormalDistXL# = CDbl(genrand_signedInt32&)
    If getRndNormalDistXL# < 0# Then
        getRndNormalDistXL# = WorksheetFunction.Norm_Inv#( _
            (getRndNormalDistXL# + Bit32Dbl#) * Bit32DblLessOneInv#, _
            mean#, Abs(std#))
    Else
        getRndNormalDistXL# = WorksheetFunction.Norm_Inv#( _
            getRndNormalDistXL# * Bit32DblLessOneInv#, mean#, Abs(std#))
    End If
End Function    'getRndNormalDistXL
'******************************************************************************

' The following C macros that used the #define keyword in the original code
'  were redefined as VBA functions below:
'******************************************************************************

' #define TEMPERING_SHIFT_U(y)  (y >> 11)
Private Function rShiftBy11(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Right Shift by 11 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for the sign bit, then
    '  divide the result by (2^11).
    rShiftBy11& = (value& And Bits11To30&) / Bit11&
    '**************************************************************************
    ' "Save the sign bit" and move it to the 20th bit position.
    If value& < 0& Then
        rShiftBy11& = rShiftBy11& Or Bit20&
    End If
End Function    'rShiftBy11
'******************************************************************************

' #define TEMPERING_SHIFT_S(y)  (y << 7)
Private Function lShiftBy7(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 7 bits.
'******************************************************************************
    ' Avoid overflow errors by preemptively masking Bits 0 through 23 before
    '  left shifting by 7 bits (Bit 24 is not masked because it will be
    '  shifted to the 31st bit of the Result, which is its Sign bit).  Then,
    '  multiply the result by (2^7).
    lShiftBy7& = (value& And Bits0To23&) * Bit7&
    '**************************************************************************
    ' Restore what was the 24th bit in Value to the 31st bit of the result.
    If (value& And Bit24&) <> 0& Then
        lShiftBy7& = lShiftBy7& Or UPPER_MASK&
    End If
End Function    'lShiftBy7
'******************************************************************************

' #define TEMPERING_SHIFT_T(y)  (y << 15)
Private Function lShiftBy15(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 15 bits.
'******************************************************************************
    ' Avoid overflow errors by preemptively masking Bits 0 through 15 before
    '  left shifting by 15 bits (Bit 16 is not masked because it will be
    '  shifted to the 31st bit of the Result, which is its Sign bit).  Then,
    '  multiply the result by (2^15).
    lShiftBy15& = (value& And Bits0To15&) * Bit15&
    '**************************************************************************
    ' Restore what was the 16th bit in Value to the 31st bit of the result.
    If (value& And Bit16&) <> 0& Then
        lShiftBy15& = lShiftBy15& Or UPPER_MASK&
    End If
End Function    'lShiftBy15
'******************************************************************************

' #define TEMPERING_SHIFT_L(y)  (y >> 18)
Private Function rShiftBy18(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Right Shift by 18 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for the sign bit, then
    '  divide the result by (2^18).
    rShiftBy18& = (value& And Bits18To30&) / Bit18&
    '**************************************************************************
    ' "Save the sign bit" and move it to the 13th bit position.
    If value& < 0& Then
        rShiftBy18& = rShiftBy18& Or Bit13&
    End If
End Function    'rShiftBy18
'******************************************************************************

' The functions below were not part of the original C code.  They are needed
'  here in order to emulate various functionalities.  They are NOT INTENDED
'  for general use!

' Some General Notes: Errors (such as positive or negative overflow) involving
'  signed number arithmetic are avoided by not using arithmetic operations
'  altogether and only relying on bitwise operations.  Thus, a Signed Value
'  representation can be treated as Unsigned for the time being, without
'  actually converting a Signed Value into an Unsigned Value.
' In the case of left bit shifts, preemptively masking only the bits that will
'  be preserved in the operation will help avoid any overflow errors.
' Note that Unsigned Right Shifts are the same as Logical Right Shifts on
'  Signed Values, where zeroes are shifted in.  Thus, the sign bit is
'  "discarded" from the original value if the number was a Signed Integer.
'******************************************************************************

Private Function lShiftBy1(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 1 bit.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 30 (this is the
    '  sign bit of the result), then multiply by (2^1).
    lShiftBy1& = (value& And Bits0To29&) * Bit1&
    '**************************************************************************
    ' Restore what was the 30th bit in Value to the 31st bit of the result.
    If (value& And Bit30&) <> 0& Then
        lShiftBy1& = lShiftBy1& Or UPPER_MASK&
    End If
End Function    'lShiftBy1

Private Function adder(ByVal num1 As Long, ByVal num2 As Long) As Long
'******************************************************************************
' This function emulates Unsigned Binary Addition (with the Carry Bit from
'  the addition of the Most Significant Bits automatically discarded) on two
'  32-bit Signed Long Integer numbers without actually converting them into
'  their corresponding Unsigned Value representations.
' This is only used during the initialization of the array "MT(0 To 623)".
' The function as written runs slightly faster for the case where "num2" is
'  zero as opposed to non-zero, regardless of the value of "num1".
'******************************************************************************
    Dim carry_n As Long, shift_n As Long
    ' "Initialize" the local variables.
    adder& = num1&
    shift_n& = num2&
    '**************************************************************************
    ' Emulate Unsigned Binary Addition in three repeating steps.
    Do While shift_n& <> 0&
        ' Generate The Carry Bits.
        carry_n& = adder& And shift_n&
        ' "Add" the two numbers.
        adder& = adder& Xor shift_n&
        ' Left Shift the carry bits by one.  Enough zeroes should be shifted
        '  into "shift_n" such that the loop will eventually terminate.
        shift_n& = lShiftBy1&(carry_n&)
    Loop
End Function    'adder
'******************************************************************************

Private Function rShiftBy30(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Right Shift by 30 bits.
'******************************************************************************
    ' Move the 30th bit to the zeroth bit position.
    If (value& And Bit30&) <> 0& Then
        rShiftBy30& = Bit0&
    Else
        rShiftBy30& = 0&
    End If
    '**************************************************************************
    ' "Save the sign bit" and move it to the first bit position.
    If value& < 0& Then
        rShiftBy30& = rShiftBy30& Or Bit1&
    End If
End Function    'rShiftBy30
'******************************************************************************

Private Function rShiftBy6(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Right Shift by 6 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for the sign bit, then
    '  divide the result by (2^6).
    rShiftBy6& = (value& And Bits6To30&) / Bit6&
    '**************************************************************************
    ' "Save the sign bit" and move it to the 25th bit position.
    If value& < 0& Then
        rShiftBy6& = rShiftBy6& Or Bit25&
    End If
End Function    'rShiftBy6
'******************************************************************************

Private Function rShiftBy5(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Right Shift by 5 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for the sign bit, then
    '  divide the result by (2^5).
    rShiftBy5& = (value& And Bits5To30&) / Bit5&
    '**************************************************************************
    ' "Save the sign bit" and move it to the 26th bit position.
    If value& < 0& Then
        rShiftBy5& = rShiftBy5& Or Bit26&
    End If
End Function    'rShiftBy5
'******************************************************************************

Private Function rShiftBy1(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Right Shift by 1 bit.
'******************************************************************************
    ' Mask all bits that will be preserved except for the sign bit, then
    '  divide the result by (2^1).
    rShiftBy1& = (value& And Bits1To30&) / Bit1&
    '**************************************************************************
    ' "Save the sign bit" and move it to the 30th bit position.
    If value& < 0& Then
        rShiftBy1& = rShiftBy1& Or Bit30&
    End If
End Function    'rShiftBy1
'******************************************************************************

Private Function lShiftBy2(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 2 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 29 (this is the
    '  sign bit of the result), then multiply by (2^2).
    lShiftBy2& = (value& And Bits0To28&) * Bit2&
    '**************************************************************************
    ' Restore what was the 29th bit in Value to the 31st bit of the result.
    If (value& And Bit29&) <> 0& Then
        lShiftBy2& = lShiftBy2& Or UPPER_MASK&
    End If
End Function    'lShiftBy2
'******************************************************************************

Private Function lShiftBy3(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 3 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 28 (this is the
    '  sign bit of the result), then multiply by (2^3).
    lShiftBy3& = (value& And Bits0To27&) * Bit3&
    '**************************************************************************
    ' Restore what was the 28th bit in Value to the 31st bit of the result.
    If (value& And Bit28&) <> 0& Then
        lShiftBy3& = lShiftBy3& Or UPPER_MASK&
    End If
End Function    'lShiftBy3
'******************************************************************************

Private Function lShiftBy5(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 5 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 26 (this is the
    '  sign bit of the result), then multiply by (2^5).
    lShiftBy5& = (value& And Bits0To25&) * Bit5&
    '**************************************************************************
    ' Restore what was the 26th bit in Value to the 31st bit of the result.
    If (value& And Bit26&) <> 0& Then
        lShiftBy5& = lShiftBy5& Or UPPER_MASK&
    End If
End Function    'lShiftBy5
'******************************************************************************

Private Function lShiftBy6(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 6 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 25 (this is the
    '  sign bit of the result), then multiply by (2^6).
    lShiftBy6& = (value& And Bits0To24&) * Bit6&
    '**************************************************************************
    ' Restore what was the 25th bit in Value to the 31st bit of the result.
    If (value& And Bit25&) <> 0& Then
        lShiftBy6& = lShiftBy6& Or UPPER_MASK&
    End If
End Function    'lShiftBy6
'******************************************************************************

' Note that "lShiftBy7()" is congruent to "TEMPERING_SHIFT_S()", which was in
'  the original C code and thus has already been defined above.

Private Function lShiftBy8(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 8 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 23 (this is the
    '  sign bit of the result), then multiply by (2^8).
    lShiftBy8& = (value& And Bits0To22&) * Bit8&
    '**************************************************************************
    ' Restore what was the 23rd bit in Value to the 31st bit of the result.
    If (value& And Bit23&) <> 0& Then
        lShiftBy8& = lShiftBy8& Or UPPER_MASK&
    End If
End Function    'lShiftBy8
'******************************************************************************

Private Function lShiftBy9(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 9 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 22 (this is the
    '  sign bit of the result), then multiply by (2^9).
    lShiftBy9& = (value& And Bits0To21&) * Bit9&
    '**************************************************************************
    ' Restore what was the 22rd bit in Value to the 31st bit of the result.
    If (value& And Bit22&) <> 0& Then
        lShiftBy9& = lShiftBy9& Or UPPER_MASK&
    End If
End Function    'lShiftBy9
'******************************************************************************

Private Function lShiftBy10(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 10 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 21 (this is the
    '  sign bit of the result), then multiply by (2^10).
    lShiftBy10& = (value& And Bits0To20&) * Bit10&
    '**************************************************************************
    ' Restore what was the 21st bit in Value to the 31st bit of the result.
    If (value& And Bit21&) <> 0& Then
        lShiftBy10& = lShiftBy10& Or UPPER_MASK&
    End If
End Function    'lShiftBy10
'******************************************************************************

Private Function lShiftBy11(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 11 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 20 (this is the
    '  sign bit of the result), then multiply by (2^11).
    lShiftBy11& = (value& And Bits0To19&) * Bit11&
    '**************************************************************************
    ' Restore what was the 20th bit in Value to the 31st bit of the result.
    If (value& And Bit20&) <> 0& Then
        lShiftBy11& = lShiftBy11& Or UPPER_MASK&
    End If
End Function    'lShiftBy11
'******************************************************************************

Private Function lShiftBy13(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 13 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 18 (this is the
    '  sign bit of the result), then multiply by (2^13).
    lShiftBy13& = (value& And Bits0To17&) * Bit13&
    '**************************************************************************
    ' Restore what was the 18th bit in Value to the 31st bit of the result.
    If (value& And Bit18&) <> 0& Then
        lShiftBy13& = lShiftBy13& Or UPPER_MASK&
    End If
End Function    'lShiftBy13
'******************************************************************************

Private Function lShiftBy14(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 14 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 17 (this is the
    '  sign bit of the result), then multiply by (2^14).
    lShiftBy14& = (value& And Bits0To16&) * Bit14&
    '**************************************************************************
    ' Restore what was the 17th bit in Value to the 31st bit of the result.
    If (value& And Bit17&) <> 0& Then
        lShiftBy14& = lShiftBy14& Or UPPER_MASK&
    End If
End Function    'lShiftBy14
'******************************************************************************

' Note that "lShiftBy15()" is congruent to "TEMPERING_SHIFT_T()", which was in
'  the original C code and thus has already been defined above.

Private Function lShiftBy16(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 16 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 15 (this is the
    '  sign bit of the result), then multiply by (2^16).
    lShiftBy16& = (value& And Bits0To14&) * Bit16&
    '**************************************************************************
    ' Restore what was the 15th bit in Value to the 31st bit of the result.
    If (value& And Bit15&) <> 0& Then
        lShiftBy16& = lShiftBy16& Or UPPER_MASK&
    End If
End Function    'lShiftBy16
'******************************************************************************

Private Function lShiftBy17(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 17 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 14 (this is the
    '  sign bit of the result), then multiply by (2^17).
    lShiftBy17& = (value& And Bits0To13&) * Bit17&
    '**************************************************************************
    ' Restore what was the 14th bit in Value to the 31st bit of the result.
    If (value& And Bit14&) <> 0& Then
        lShiftBy17& = lShiftBy17& Or UPPER_MASK&
    End If
End Function    'lShiftBy17
'******************************************************************************

Private Function lShiftBy18(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 18 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 13 (this is the
    '  sign bit of the result), then multiply by (2^18).
    lShiftBy18& = (value& And Bits0To12&) * Bit18&
    '**************************************************************************
    ' Restore what was the 13th bit in Value to the 31st bit of the result.
    If (value& And Bit13&) <> 0& Then
        lShiftBy18& = lShiftBy18& Or UPPER_MASK&
    End If
End Function    'lShiftBy18
'******************************************************************************

Private Function lShiftBy19(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 19 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 12 (this is the
    '  sign bit of the result), then multiply by (2^19).
    lShiftBy19& = (value& And Bits0To11&) * Bit19&
    '**************************************************************************
    ' Restore what was the 12th bit in Value to the 31st bit of the result.
    If (value& And Bit12&) <> 0& Then
        lShiftBy19& = lShiftBy19& Or UPPER_MASK&
    End If
End Function    'lShiftBy19
'******************************************************************************

Private Function lShiftBy20(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 20 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 11 (this is the
    '  sign bit of the result), then multiply by (2^20).
    lShiftBy20& = (value& And Bits0To10&) * Bit20&
    '**************************************************************************
    ' Restore what was the 12th bit in Value to the 31st bit of the result.
    If (value& And Bit11&) <> 0& Then
        lShiftBy20& = lShiftBy20& Or UPPER_MASK&
    End If
End Function    'lShiftBy20
'******************************************************************************

Private Function lShiftBy22(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 22 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 22 (this is the
    '  sign bit of the result), then multiply by (2^9).
    lShiftBy22& = (value& And Bits0To8&) * Bit22&
    '**************************************************************************
    ' Restore what was the 9th bit in Value to the 31st bit of the result.
    If (value& And Bit9&) <> 0& Then
        lShiftBy22& = lShiftBy22& Or UPPER_MASK&
    End If
End Function    'lShiftBy22
'******************************************************************************

Private Function lShiftBy24(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 24 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 7 (this is the
    '  sign bit of the result), then multiply by (2^24).
    lShiftBy24& = (value& And Bits0To6&) * Bit24&
    '**************************************************************************
    ' Restore what was the 7th bit in Value to the 31st bit of the result.
    If (value& And Bit7&) <> 0& Then
        lShiftBy24& = lShiftBy24& Or UPPER_MASK&
    End If
End Function    'lShiftBy24
'******************************************************************************

Private Function lShiftBy26(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 26 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 5 (this is the
    '  sign bit of the result), then multiply by (2^26).
    lShiftBy26& = (value& And Bits0To4&) * Bit26&
    '**************************************************************************
    ' Restore what was the 5th bit in Value to the 31st bit of the result.
    If (value& And Bit5&) <> 0& Then
        lShiftBy26& = lShiftBy26& Or UPPER_MASK&
    End If
End Function    'lShiftBy26
'******************************************************************************

Private Function lShiftBy27(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 27 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 4 (this is the
    '  sign bit of the result), then multiply by (2^27).
    lShiftBy27& = (value& And Bits0To3&) * Bit27&
    '**************************************************************************
    ' Restore what was the 4th bit in Value to the 31st bit of the result.
    If (value& And Bit4&) <> 0& Then
        lShiftBy27& = lShiftBy27& Or UPPER_MASK&
    End If
End Function    'lShiftBy27
'******************************************************************************

Private Function lShiftBy28(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 28 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 3 (this is the
    '  sign bit of the result), then multiply by (2^28).
    lShiftBy28& = (value& And Bits0To2&) * Bit28&
    '**************************************************************************
    ' Restore what was the 3rd bit in Value to the 31st bit of the result.
    If (value& And Bit3&) <> 0& Then
        lShiftBy28& = lShiftBy28& Or UPPER_MASK&
    End If
End Function    'lShiftBy28
'******************************************************************************

Private Function lShiftBy29(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 29 bits.
'******************************************************************************
    ' Mask all bits that will be preserved except for Bit 2 (this is the
    '  sign bit of the result), then multiply by (2^29).
    lShiftBy29& = (value& And Bits0To1&) * Bit29&
    '**************************************************************************
    ' Restore what was the 2nd bit in Value to the 31st bit of the result.
    If (value& And Bit2&) <> 0& Then
        lShiftBy29& = lShiftBy29& Or UPPER_MASK&
    End If
End Function    'lShiftBy29
'******************************************************************************

Private Function lShiftBy30(ByVal value As Long) As Long
'******************************************************************************
' This Function emulates an Unsigned Left Shift by 30 bits.
'******************************************************************************
    ' Move the zeroth bit to the 30th bit position.
    If (value& And Bit0&) <> 0& Then
        lShiftBy30& = Bit30&
    Else
        lShiftBy30& = 0&
    End If
    '**************************************************************************
    ' Move the first bit to the "sign bit position".
    If (value& And Bit1&) <> 0& Then
        lShiftBy30& = lShiftBy30& Or UPPER_MASK&
    End If
End Function    'lShiftBy30
'******************************************************************************

'******************************************************************************
' A copy of the original C code has been provided below for reference:
'******************************************************************************
' /*
' A C-program for MT19937, with initialization improved 2002/1/26.
' Coded by Takuji Nishimura and Makoto Matsumoto.
'
' Before using, initialize the state by using init_genrand(seed)
' or init_by_array(init_key, key_length).
'
' Copyright (C) 1997 - 2002, Makoto Matsumoto and Takuji Nishimura,
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions
' are met:
'
'   1. Redistributions of source code must retain the above copyright
'      notice, this list of conditions and the following disclaimer.
'
'   2. Redistributions in binary form must reproduce the above copyright
'      notice, this list of conditions and the following disclaimer in the
'      documentation and/or other materials provided with the distribution.
'
'   3. The names of its contributors may not be used to endorse or promote
'      products derived from this software without specific prior written
'      permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
' "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
' LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
' A PARTICULAR PURPOSE ARE DISCLAIMED.  IN NO EVENT SHALL THE COPYRIGHT OWNER
' OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
' EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
' PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
' PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
' LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
' NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'
' Any feedback is very welcome.
' http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/emt.html
' email: m-mat @ math.sci.hiroshima-u.ac.jp (remove space)
' */
'
' #include <stdio.h>
'
' /* Period parameters */
' #define N 624
' #define M 397
' #define MATRIX_A 0x9908b0dfUL   /* constant vector a */
' #define UPPER_MASK 0x80000000UL /* most significant w-r bits */
' #define LOWER_MASK 0x7fffffffUL /* least significant r bits */
'
' static unsigned long mt[N]; /* the array for the state vector  */
' static int mti=N+1; /* mti==N+1 means mt[N] is not initialized */
'
' /* initializes mt[N] with a seed */
' void init_genrand(unsigned long s)
' {
'     mt[0]= s & 0xffffffffUL;
'     for (mti=1; mti<N; mti++) {
'         mt[mti] =
'         (1812433253UL * (mt[mti-1] ^ (mt[mti-1] >> 30)) + mti);
'         /* See Knuth TAOCP Vol2. 3rd Ed. P.106 for multiplier. */
'         /* In the previous versions, MSBs of the seed affect   */
'         /* only MSBs of the array mt[].                        */
'         /* 2002/01/09 modified by Makoto Matsumoto             */
'         mt[mti] &= 0xffffffffUL;
'         /* for >32 bit machines */
'     }
' }
'
' /* initialize by an array with array-length */
' /* init_key is the array for initializing keys */
' /* key_length is its length */
' /* slight change for C++, 2004/2/26 */
' void init_by_array(unsigned long init_key[], int key_length)
' {
'     int i, j, k;
'     init_genrand(19650218UL);
'     i=1; j=0;
'     k = (N>key_length ? N : key_length);
'     for (; k; k--) {
'         mt[i] = (mt[i] ^ ((mt[i-1] ^ (mt[i-1] >> 30)) * 1664525UL))
'           + init_key[j] + j; /* non linear */
'         mt[i] &= 0xffffffffUL; /* for WORDSIZE > 32 machines */
'         i++; j++;
'         if (i>=N) { mt[0] = mt[N-1]; i=1; }
'         if (j>=key_length) j=0;
'     }
'     for (k=N-1; k; k--) {
'         mt[i] = (mt[i] ^ ((mt[i-1] ^ (mt[i-1] >> 30)) * 1566083941UL))
'           - i; /* non linear */
'         mt[i] &= 0xffffffffUL; /* for WORDSIZE > 32 machines */
'         i++;
'         if (i>=N) { mt[0] = mt[N-1]; i=1; }
'     }
'
'     mt[0] = 0x80000000UL; /* MSB is 1; assuring non-zero initial array */
' }
'
' /* generates a random number on [0,0xffffffff]-interval */
' unsigned long genrand_int32(void)
' {
'     unsigned long y;
'     static unsigned long mag01[2]={0x0UL, MATRIX_A};
'     /* mag01[x] = x * MATRIX_A  for x=0,1 */
'
'     if (mti >= N) { /* generate N words at one time */
'         int kk;
'
'         if (mti == N+1)   /* if init_genrand() has not been called, */
'             init_genrand(5489UL); /* a default initial seed is used */
'
'         for (kk=0;kk<N-M;kk++) {
'             y = (mt[kk]&UPPER_MASK)|(mt[kk+1]&LOWER_MASK);
'             mt[kk] = mt[kk+M] ^ (y >> 1) ^ mag01[y & 0x1UL];
'         }
'         for (;kk<N-1;kk++) {
'             y = (mt[kk]&UPPER_MASK)|(mt[kk+1]&LOWER_MASK);
'             mt[kk] = mt[kk+(M-N)] ^ (y >> 1) ^ mag01[y & 0x1UL];
'         }
'         y = (mt[N-1]&UPPER_MASK)|(mt[0]&LOWER_MASK);
'         mt[N-1] = mt[M-1] ^ (y >> 1) ^ mag01[y & 0x1UL];
'
'         mti = 0;
'     }
'
'     y = mt[mti++];
'
'     /* Tempering */
'     y ^= (y >> 11);
'     y ^= (y << 7) & 0x9d2c5680UL;
'     y ^= (y << 15) & 0xefc60000UL;
'     y ^= (y >> 18);
'
'     return y;
' }
'
' /* generates a random number on [0,0x7fffffff]-interval */
' long genrand_int31(void)
' {
'     return (long)(genrand_int32()>>1);
' }
'
' /* generates a random number on [0,1]-real-interval */
' double genrand_real1(void)
' {
'     return genrand_int32()*(1.0/4294967295.0);
'     /* divided by 2^32-1 */
' }
'
' /* generates a random number on [0,1)-real-interval */
' double genrand_real2(void)
' {
'     return genrand_int32()*(1.0/4294967296.0);
'     /* divided by 2^32 */
' }
'
' /* generates a random number on (0,1)-real-interval */
' double genrand_real3(void)
' {
'     return (((double)genrand_int32()) + 0.5)*(1.0/4294967296.0);
'     /* divided by 2^32 */
' }
'
' /* generates a random number on [0,1) with 53-bit resolution*/
' double genrand_res53(void)
' {
'     unsigned long a=genrand_int32()>>5, b=genrand_int32()>>6;
'     return(a*67108864.0+b)*(1.0/9007199254740992.0);
' }
' /* These real versions are due to Isaku Wada, 2002/01/09 added */
'
' int main(void)
' {
'     int i;
'     unsigned long init[4]={0x123, 0x234, 0x345, 0x456}, length=4;
'     init_by_array(init, length);
'     printf("1000 outputs of genrand_int32()\n");
'     for (i=0; i<1000; i++) {
'       printf("%10lu ", genrand_int32());
'       if (i%5==4) printf("\n");
'     }
'     printf("\n1000 outputs of genrand_real2()\n");
'     for (i=0; i<1000; i++) {
'       printf("%10.8f ", genrand_real2());
'       if (i%5==4) printf("\n");
'     }
'     return 0;
' }
'******************************************************************************
' EOF
'******************************************************************************

' End Of Module



