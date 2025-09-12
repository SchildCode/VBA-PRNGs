Option Explicit
Private Const BIT31^ = 2 ^ 31
Private Const BIT32^ = 2 ^ 32
Private Const DIV32# = 2.3283064365387E-10 '=1/(2^32)
'Routines for testing PRNGs

Private Sub PRNG_basic_test()
    Dim i&
    Dim U32^
    Dim s^
    Dim x#, xMin#, xMean#, xMax#, t0#
'---
    'Initialize & seed, if needed:
    '<=====================
    'Call SFC32_init(1234^)
    Call PCG32_init(1234^)
    's^ = 1 'start for mix32
    '=====================>
'---
    xMin = 1#
    t0 = Timer
    For i = 1 To 10000000
        'Generare new PRNG:
        '<=====================
        'U32 = rnd_U32()
        'U32 = RAND_U32()
        'U32 = RANDARRAY_U32()
        'U32 = fmix32_U32(s)
        'U32 = mix32_U32(s)
        'U32 = LFIB4_U32()
        'U32 = MTran_U32()
        'U32 = SFC32_U32()
        'U32 = Bryc32_U32()
        U32 = PCG32_U32()
        '=====================>

        'x = U32 * DIV32
        'If x < xMin Then xMin = x Else If xMax < x Then xMax = x
        'xMean = xMean + (x - xMean) / i
    Next
    Debug.Print "xMin = " & xMin, "xMean = " & xMean, "xMax = " & xMax, "t = " & Timer - t0 '414
End Sub

Public Sub GenerateENTFile()
    'Generate binary file for testing on the online service at http://www.cacert.at/random/
    'Ensure that the tested pseudorandom number generator is configured to output a 8-byte LongLon containing a 32-bit unsigned integer, i.e. 0 to 2^32-1
    Const FILE_PATH As String = "C:\msys64\home\pgs\ENT_test.bin"
    Const TARGET_SIZE_MB As Long = 256 'Mb
    Const BYTES_PER_CHUNK As Long = 4 'Each PRNG call outputs a 32-bit double (4 bytes)
    Const BUFFER_SIZE_BYTES As Long = 1048576 ' 1 MB buffer size
    Const BUFFER_SIZE_PRNG_VALUES As Long = BUFFER_SIZE_BYTES / BYTES_PER_CHUNK ' Number of PRNG values per buffer
    'Variables
    Dim FileNumber As Integer
    Dim TotalBytes&
    Dim i&
    Dim U32^
    Dim i32& 'Double containing an unsigned 4-byte random number 0 to +2^32-1 inclusive, i.e. [0, 2^32-1]
    Dim iMin#, iMean#, iMax#, iNum&
    Dim t0#
    Dim buf&(0 To BUFFER_SIZE_PRNG_VALUES - 1) 'Must be 4-byte Long
    Dim bytesWritten&
'---
    TotalBytes = TARGET_SIZE_MB * 1024 * 1024 'Calculate total bytes to generate
    FileNumber = FreeFile 'Get a free file number
    Open FILE_PATH For Binary Access Write As #FileNumber 'Open the file for binary output. This creates a new file or overwrites an existing one.
    Debug.Print "Generating a binary file of approx. " & TARGET_SIZE_MB & " MB in steps of " & BUFFER_SIZE_BYTES / 1024 & " KB chunks."

    'Initialize & seed, if needed:
    '<===================== Uncomment only line for the relevant PRNG
    Call SFC32_init(1234^)
    'Call PCG32_init(1234^)
    '=====================>

    bytesWritten = 0
    iNum = 0
    iMin = 2 ^ 31
    t0 = Timer
    Do While bytesWritten < TotalBytes
        'Fill the buffer with PRNG values
        For i = 0 To BUFFER_SIZE_PRNG_VALUES - 1
            iNum = iNum + 1

            'Generate an 8-byte LongLong containing a 4-byte unsigned [0, 2^32-1]
            '<============================= Uncomment only line for the relevant PRNG
            U32 = Bryc32_U32()
            '=============================>

            If U32 < BIT31 Then i32 = CLng(U32) Else i32 = CLng(U32 - BIT32) 'convert to 4-byte Long [-2^31 +2^31], preserving bit order, for testing on TestU01
            buf(i) = i32
            If i32 < iMin Then iMin = i32 Else If iMax < i32 Then iMax = i32
            iMean = iMean + (i32 - iMean) / iNum
        Next i
        Put #FileNumber, , buf
        bytesWritten = bytesWritten + BUFFER_SIZE_BYTES
        If bytesWritten Mod (10 * BUFFER_SIZE_BYTES) = 0 Then Debug.Print "Wrote " & bytesWritten * 100# / TotalBytes & "%"
        DoEvents
    Loop
    Debug.Print "Output " & TotalBytes & " bytes in " & Timer - t0 & " seconds, speed " & TotalBytes / (Timer - t0) & " Bytes/second." 'With the in-loop debug commented out"
    Debug.Print "Min=" & iMin, "Mean=" & iMean, "Max=" & iMax

    ' --- Clean up ---
    Close #FileNumber
    Debug.Print "File generation complete: " & FILE_PATH
    Exit Sub
    
ErrorHandler:
    ' Handle potential errors, such as file access issues
    MsgBox "An error occurred: " & Err.Description, vbCritical, "File Generation Error"
    If FileNumber <> 0 Then Close #FileNumber
End Sub

Sub StreamToCrushViaPipe()
    ' - For continuous piped input from any Excel VBA random number generator to TestU01 (SmallCrush, Crush or BigCrush), avoiding the need to save intermediate file on disk.
    ' - Wiki into on TestU01: https://en.wikipedia.org/wiki/TestU01
    ' - TestU01 webside: https://simul.iro.umontreal.ca/testu01/tu01.html
    '
    ' INSTRUCTIONS FOR USE:
    ' 1. Install MSYS from https://www.msys2.org/
    ' 2. Start MSYS2 MINGW64 terminal window
    ' 3. In the MINGW64 terminal window, install TestU01 package, with bash command-line "pacman -S mingw-w64-x86_64-testu01"
    '    More info: https://packages.msys2.org/packages/mingw-w64-x86_64-testu01
    ' 4. Put this c-code file in folder C:\msys64\home\<username>
    ' 5. Choose between SmallCrush, Crush or BugCrush test batteries by uncommenting the relevant line in the c-code below. Save this file.
    ' 6. In the MINGW64 terminal window, build crush_from_pipe.exe with the bash command-line:
    '    "gcc -O3 -march=native -pipe -s crush_from_pipe.c -o crush_from_pipe -ltestu01 -lprobdist -lm"
    ' 7. Open the pipe by bash command-line "./crush_from_pipe | tee results.txt"
    '    Note: It then prints “Waiting for VBA writer...”. The tee saves all TestU01 output to file results.txt
    ' 8. In Excel VBA IDE, run subroutine "StreamToCrushViaPipe". VBA then generates a stream of PRNG 4-byte numbers to data via the pipe.
    '    Note: The data stream is not written to disk but is stream in memory (with small OS buffers) flowing from the VBA process to this C harness.
    '    Note: The pipe \\.\pipe\RNGStream lies in the kernel’s Named Pipe File System (NPFS), not on physical NTFS disk.
    ' 9. When Crush finishes, it closes the pipe; VBA subroutine "StreamToCrushViaPipe" then terminates when it detects that the pipe is closed.
    ' 10. The test battery results are in file results.txt
    ' AUTHOR: Peter.Schild@OsloMet.no
        
    Const PIPE$ = "\\.\pipe\RNGStream" 'Excatly same name as given in C source code crush_from_pipe.c
    Const Chunk& = 262144 'Numbers per batch (this number has been tuned for speed)
    Const oneGb# = 1024# * 1024# * 1024# 'bits/Gb
    Const SESSION_BYTES# = 1.5 * oneGb '~1.5 Gb per session (just below the 2 Gb data limit per session in Windows)
    Const batchesPerSession& = (SESSION_BYTES / (Chunk * 4&)) 'Number of chunks output to keep within the 2 Gb limit per file output session

    Dim fh%
    Dim buf&(1 To Chunk)
    Dim i&, j&, Nretry&
    Dim s^, U32^ 'LongLong containing unsigned 32-bit number in range [0, 2^32-1]
    Dim i32&, iMax&, iMin&, iMean#, iNum& 'i32 is a signed Long containing a signed 32-bit integer in range [-2^31, 2^31-1]
    Dim t0 As Date
'---
    'Initialize & seed, if needed:
    '<===================== Uncomment only line for the relevant PRNG
    Call SFC32_init(1234^)
    'Call PCG32_init(1234^)
    '=====================>

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    t0 = Now()
    Debug.Print
    Debug.Print "Started at " & t0 & " ..."

    Do 'reapeated sessions of 1.5 Gb data ad infinitum until the C-harness closes the port
        fh = FreeFile
        Do 'retry open file
            On Error Resume Next
            Open PIPE For Binary Access Write As #fh 'Connect to the pipe (the C harness must already be waiting)
            If Err.Number <> 0 Then
                Nretry = Nretry + 1
                If 9 < Nretry Then GoTo done 'Give up after the 10th attempt. This usually never happens during a Crush test, but can if the PC is too busy doing other stuff
                Debug.Print "Restart at N=" & iNum & " at " & Now()
                Application.Wait Now + #12:00:02 AM# 'Wait 2 seconds before retrying to open PIPE
            Else 'success
                Exit Do
            End If
        Loop
        For i = 1 To batchesPerSession
            For j = 1 To Chunk
            
                'Generate an 8-byte LongLong containing a 4-byte unsigned [0, 2^32-1]
                '<============================= Uncomment only line for the relevant PRNG
                'U32 = rnd_U32()
                'U32 = RAND_U32()
                'U32 = RANDARRAY_U32()
                'U32 = fmix32_U32(s)
                '32 = mix32_U32(s)
                'U32 = LFIB4_U32()
                'U32 = MTran_U32()
                'U32 = SFC32_U32()
                U32 = Bryc32_U32()
                'U32 = PCG32_U32()
                '=============================>

                If U32 < BIT31 Then i32 = CLng(U32) Else i32 = CLng(U32 - BIT32) 'convert to 4-byte Long [-2^31 +2^31], preserving bit order, for testing on TestU01
                buf(j) = i32
                iNum = iNum + 1
                iMean = iMean + (CDbl(i32) - iMean) / iNum 'running mean, just for curiosity
                If i32 < iMin Then iMin = i32 Else If iMax < i32 Then iMax = i32 'max and min values, just for curiosity
            Next
            On Error Resume Next
            Put #fh, , buf 'Write the whole buffer (4*CHUNK bytes) in one go, for speed
            If Err.Number <> 0 Then
                Debug.Print "Write failed. Port probably closed."
                GoTo done
            End If
            DoEvents
        Next
        Close #fh 'end this 1.5 Gb session, time to start a new one
        DoEvents
        Application.Wait Now + #12:00:02 AM# 'Wait 2 seconds to ensure C-harness is ready to listen
    Loop

done:
    On Error Resume Next
    Close #fh
    
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
    
    Debug.Print "Pipe closed at " & Now() & ", " & Round((Now() - t0) * 24 * 60, 2) & " minutes (" & Round((Now() - t0) * 24, 1) & " hours)."
    Debug.Print "Generated " & iNum & " random numbers, approx 2^" & Round(Log(iNum) / Log(2#), 1) & " (" & Round(iNum * 4# / oneGb, 3) & " Gb)"
    Debug.Print "Basic statistics for random signed 4-byte Long: min=" & iMin & ", mean=" & iMean & ", max=" & iMax
    Beep
End Sub
