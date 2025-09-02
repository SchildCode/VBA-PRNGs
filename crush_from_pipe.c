// crush_from_pipe.c
// 1. Build (MSYS2 MinGW64):  gcc -O3 -march=native -pipe -s crush_from_pipe.c -o crush_from_pipe -ltestu01 -lprobdist -lm
// 2. Run: ./crush_from_pipe | tee results.txt  
// 3. Then start the VBA sender in file ShuffleSim_v2.xlsm

#include <stdio.h>
#include <stdint.h>
#include <windows.h>
#include <testu01/unif01.h>
#include <testu01/bbattery.h>

#ifndef ERROR_BROKEN_PIPE
#define ERROR_BROKEN_PIPE 109
#endif

#define PIPE_NAME "\\\\.\\pipe\\RNGStream"
#define BUF_WORDS (1u<<20)  // buffer: 1,048,576 32-bit words

static HANDLE hPipe = INVALID_HANDLE_VALUE;
static uint32_t buf[BUF_WORDS];
static size_t   pos = 0, endw = 0;
static unsigned long long words_read = 0;

static void die(const char *msg) {
    fprintf(stderr, "FATAL: %s (GetLastError=%lu)\n", msg, (unsigned long)GetLastError());
    if (hPipe != INVALID_HANDLE_VALUE) CloseHandle(hPipe);
    exit(1);
}

static void reconnect_client(void); // forward

static void fill_buffer_blocking(void) {
    pos = 0; 
	endw = 0;
    uint8_t *p = (uint8_t*)buf;
    const DWORD want = (DWORD)sizeof(buf);
    DWORD got_total = 0;

    while (got_total < want) {
        DWORD got = 0;
        BOOL ok = ReadFile(hPipe, p + got_total, want - got_total, &got, NULL);
        if (!ok) {
            DWORD e = GetLastError();
            if (e == ERROR_BROKEN_PIPE) {         // writer closed: accept next session
                reconnect_client();
                continue;                          // keep filling the remainder
            }
            die("ReadFile failed (unexpected)");
        }
        if (got == 0) {                            // graceful EOF -> treat like reconnect
            reconnect_client();
            continue;
        }
        got_total += got;
    }
    endw = want / 4;
    words_read += endw;
    if ((words_read & ((1ULL<<24)-1)) == 0) {
        fprintf(stderr, "streamed %llu words (%.3f Gb)\n",words_read, (words_read*4.0)/(1024.0*1024.0*1024.0));
    }
}

static void reconnect_client(void) {
    fprintf(stderr, "writer closed; waiting for next session...\n");
    if (hPipe != INVALID_HANDLE_VALUE) {
        CloseHandle(hPipe); // You can reuse the same handle via DisconnectNamedPipe, but simplest is recreate:
    }
    // (re)create and wait for the next client
    HANDLE h = CreateNamedPipeA(PIPE_NAME, PIPE_ACCESS_INBOUND, PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT, 1, 65536, 65536, 0, NULL);
    if (h == INVALID_HANDLE_VALUE) die("CreateNamedPipe (reconnect) failed");
    BOOL ok = ConnectNamedPipe(h, NULL) ? TRUE : (GetLastError() == ERROR_PIPE_CONNECTED);
    if (!ok) die("ConnectNamedPipe (reconnect) failed");
    hPipe = h;
    fprintf(stderr, "reconnected; continuing stream...\n");
}

// TestU01 callback: return 32 random bits from the pipe
static unsigned int next_bits(void) {
    if (pos == endw) fill_buffer_blocking();
    return buf[pos++];
}

static void wait_for_client(void) {
    HANDLE h = CreateNamedPipeA(PIPE_NAME, PIPE_ACCESS_INBOUND, PIPE_TYPE_BYTE | PIPE_READMODE_BYTE | PIPE_WAIT, 1, 65536, 65536, 0, NULL);
    if (h == INVALID_HANDLE_VALUE) die("CreateNamedPipe failed");

    printf("Waiting for VBA writer to connect on %s ...\n", PIPE_NAME);
    fflush(stdout);
    BOOL ok = ConnectNamedPipe(h, NULL) ? TRUE : (GetLastError() == ERROR_PIPE_CONNECTED);
    if (!ok) die("ConnectNamedPipe failed");
    hPipe = h;
    printf("Client connected. Running bbattery_Crush...\n");
    fflush(stdout);
}

int main(void) {
    wait_for_client();

    // Wrap the pipe as a generator and run Crush
    unif01_Gen *gen = unif01_CreateExternGenBits("VBA_pipe32", next_bits);
	//bbattery_SmallCrush(gen);
    bbattery_Crush(gen);
    unif01_DeleteExternGenBits(gen);

    CloseHandle(hPipe);
    puts("Done. Pipe closed.");
    return 0;
}
