Option Explicit

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
' 4. Save c-code file "crush_from_pipe.c" in folder C:\msys64\home\<username>
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
    Const CHUNK& = 262144 'Numbers per batch (this number has been tuned for speed)
    Const oneGb# = 1024# * 1024# * 1024# 'bits/Gb
    Const SESSION_BYTES# = 1.5 * oneGb '~1.5 Gb per session (just below the 2 Gb data limit per session in Windows)
    Const batchesPerSession& = (SESSION_BYTES / (CHUNK * 4&)) 'Number of chunks output to keep within the 2 Gb limit per file output session

    Dim fh%
    Dim buf&(1 To CHUNK)
    Dim i&
    Dim j&
    Dim Nretry&
    Dim xMax&, xMin&, xMean#, xNum#
    Dim t0 As Date

    'Seed or initialize the PRNG:
    '<============================= Uncomment only line for the relevant PRNG if needed
    'SplitMix32 (1)
    'i = SplitMix32_v2(1)
    'i = LFIB4b(2627395)
    'Call SFC32_Init
    'Call SFC32v2_Init
    'Call SFC32_seed(1, 1) 'Initialize with SplitMix32
    'PCG32_Init seed_lo:=12345, seed_hi:=0, seq_lo:=54, seq_hi:=0
    'Call PCG32v5_Init(1, 1)
    'Call PCG32v5_seed(1, 1) 'Initialize with SplitMix32
    '=============================>

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    xMean = 0
    xNum = 0
    xMin = 0
    xMax = 0

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
                Debug.Print "Restart at N=" & xNum & " at " & Now()
                Application.Wait Now + #12:00:02 AM# 'Wait 2 seconds before retrying to open PIPE
            Else 'success
                Exit Do
            End If
        Loop
        For i = 1 To batchesPerSession
            For j = 1 To CHUNK 'Generate 4-byte signed long [-2^31, 2^31) with identical 4-byte pattern as unsigned integer [0, 2^32), and store in buffer
                '<============================= Uncomment only line for the relevant PRNG
                'buf(j) = SplitMix32() '0.77 minutes | Passed Crush
                'buf(j) = SplitMix32_v2()
                'buf(j) = VBA_U01() 'converts from U01
                'buf(j) = LFIB4b() 'passed SmallCrush, 0.92 minutes
                'buf(j) = SFC32_NextUInt32 '0.78 minutes | Passed Crush
                'buf(j) = SFC32v2_NextUInt32 '1.2 minutes | Passed Crush
                'buf(j) = PCG32_NextUInt32() 'passed 3.37 mins identical results to PCG32v4_NextUInt32
                'buf(j) = PCG32v5_NextUInt32() 'passed in 2.23 minutes
                buf(j) = ExcelRAND(j, CHUNK)
                'buf(j) = ExcelRAND_LowBits(j, CHUNK)
                'buf(j) = ExcelRAND_HighBits(j, CHUNK)
                '=============================>
                xNum = xNum + 1
                xMean = xMean + (CDbl(buf(j)) - xMean) / xNum 'running mean, just for curiosity
                If buf(j) < xMin Then xMin = buf(j) Else If xMax < buf(j) Then xMax = buf(j) 'max and min values, just for curiosity
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
    Debug.Print "Generated " & xNum & " random numbers, approx 2^" & Round(Log(xNum) / Log(2#), 1) & " (" & Round(xNum * 4# / oneGb, 3) & " Gb)"
    Debug.Print "Basic statistics for random signed 4-byte Long: min=" & xMin & ", mean=" & xMean & ", max=" & xMax
    Beep

End Sub
