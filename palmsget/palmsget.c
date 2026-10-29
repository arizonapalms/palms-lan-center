/* palmsget.exe - downloads the Game Library's files for Palms LAN Center, on Windows 95 OSR2 and later.
 *
 *   palmsget [-c] [-nc] [-nv] [-t tries] [-T seconds] [--no-proxy] [-o logfile] [-P folder]
 *            [-O file] [-i urlfile] [URL...]
 *     -c    carry on with a file that is partly downloaded
 *     -nc   skip a file that is already there
 *     -t    tries for each file (default 20)
 *     -T    time limit in seconds for connecting and for each read (default 900)
 *     -o    write what happens to this file
 *     -P    save into this folder
 *     -O    save the (one) download as this file
 *     -i    read the addresses from this file, one a line
 * Exit codes: 0 all done, 1 something else, 2 bad options, 3 unable to write the file,
 * 4 unable to connect, 8 the server said no (404...). Build with build.sh.
 */
#include <windows.h>
#include <wininet.h>

#define VERSION "1.1"
#define CHUNK 65536
#define MAXURLS 512

static char g_buf[CHUNK];
static char g_line[1200];
static char g_folder[MAX_PATH];
static char g_path[MAX_PATH * 2];
static char g_name[MAX_PATH];
static char g_hdr[64];
static char g_req[4096];
static char g_urlfile[MAX_PATH];
static char g_logname[MAX_PATH];
static char g_outfile[MAX_PATH * 2];
static char g_cmd[32768];
static char g_list[65536];
static char *g_urls[MAXURLS];
static int g_nurls;

static HANDLE g_log = INVALID_HANDLE_VALUE;
static int g_continue, g_noclobber, g_tries = 20, g_timeout = 900, g_is9x;

/* the compiler may call these for struct copies */
void *memset(void *d, int c, unsigned int n) { unsigned char *p = d; while (n--) *p++ = (unsigned char)c; return d; }
void *memcpy(void *d, const void *s, unsigned int n) { unsigned char *a = d; const unsigned char *b = s; while (n--) *a++ = *b++; return d; }

static void put(const char *s)
{
    DWORD w;
    if (g_log != INVALID_HANDLE_VALUE) WriteFile(g_log, s, lstrlenA(s), &w, NULL);
}

static void stamp(void)
{
    SYSTEMTIME t;
    char s[32];
    GetLocalTime(&t);
    wsprintfA(s, "%04u-%02u-%02u %02u:%02u:%02u ", t.wYear, t.wMonth, t.wDay, t.wHour, t.wMinute, t.wSecond);
    put(s);
}

/* one line in the log, the time first */
static void logline(const char *s)
{
    stamp();
    put(s);
    put("\r\n");
    if (g_log != INVALID_HANDLE_VALUE) FlushFileBuffers(g_log);
}

/* the next word of a command line: "quoted words" stay together */
static char *nextword(char *p, char *out, int max)
{
    int n = 0;
    while (*p == ' ' || *p == '\t') p++;
    if (!*p) return NULL;
    while (*p && *p != ' ' && *p != '\t') {
        if (*p == '"') {
            p++;
            while (*p && *p != '"') { if (n < max - 1) out[n++] = *p; p++; }
            if (*p == '"') p++;
        } else {
            if (n < max - 1) out[n++] = *p;
            p++;
        }
    }
    out[n] = 0;
    return p;
}

static int number(const char *s)
{
    int n = 0;
    while (*s >= '0' && *s <= '9') n = n * 10 + (*s++ - '0');
    return n;
}

static int same(const char *a, const char *b) { return lstrcmpA(a, b) == 0; }

static void addurl(char *u)
{
    if (*u && g_nurls < MAXURLS) g_urls[g_nurls++] = u;
}

/* the addresses in a -i file, one a line */
static int readlist(const char *file, char **space, int room)
{
    HANDLE h;
    DWORD got = 0;
    char *p, *start;
    h = CreateFileA(file, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
    if (h == INVALID_HANDLE_VALUE) return 0;
    ReadFile(h, *space, room - 1, &got, NULL);
    CloseHandle(h);
    (*space)[got] = 0;
    p = start = *space;
    for (;;) {
        if (*p == '\r' || *p == '\n' || *p == 0) {
            int end = (*p == 0);
            *p = 0;
            while (*start == ' ' || *start == '\t') start++;
            if (*start) addurl(start);
            if (end) break;
            start = p + 1;
        }
        p++;
    }
    *space = p + 1;
    return 1;
}

static int hexval(char c)
{
    if (c >= '0' && c <= '9') return c - '0';
    if (c >= 'a' && c <= 'f') return c - 'a' + 10;
    if (c >= 'A' && c <= 'F') return c - 'A' + 10;
    return -1;
}

/* the file name for an address: its last part, with %20 and friends turned back into characters */
static void filename(const char *url, char *out, int max)
{
    const char *s = url, *last = url, *q;
    int n = 0;
    while (*s && *s != '?' && *s != '#') { if (*s == '/') last = s + 1; s++; }
    q = last;
    while (q < s && n < max - 1) {
        if (*q == '%' && q + 2 < s && hexval(q[1]) >= 0 && hexval(q[2]) >= 0) {
            out[n++] = (char)(hexval(q[1]) * 16 + hexval(q[2]));
            q += 3;
        } else {
            char c = *q++;
            if (c == '\\' || c == ':' || c == '*' || c == '?' || c == '"' || c == '<' || c == '>' || c == '|') c = '_';
            out[n++] = c;
        }
    }
    out[n] = 0;
    if (n == 0) lstrcpyA(out, "index.html");
}

/* the address as it goes to the server: spaces become %20, the way wget sends them */
static void request_url(const char *url, char *out, int max)
{
    int n = 0;
    while (*url && n < max - 4) {
        if (*url == ' ') { out[n++] = '%'; out[n++] = '2'; out[n++] = '0'; }
        else out[n++] = *url;
        url++;
    }
    out[n] = 0;
}

/* the size of a file, or -1 when it isn't there */
static long long sizeof_file(const char *path)
{
    HANDLE h;
    DWORD lo, hi = 0;
    h = CreateFileA(path, 0, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
    if (h == INVALID_HANDLE_VALUE)
        h = CreateFileA(path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
    if (h == INVALID_HANDLE_VALUE) return -1;
    SetLastError(NO_ERROR);
    lo = GetFileSize(h, &hi);
    CloseHandle(h);
    if (lo == 0xFFFFFFFF && GetLastError() != NO_ERROR) return -1;
    return ((long long)hi << 32) | lo;
}

static void numtext(long long v, char *out)
{
    char t[24];
    int n = 0, i;
    if (v <= 0) { lstrcpyA(out, "0"); return; }
    while (v > 0 && n < 22) { t[n++] = (char)('0' + v % 10); v /= 10; }
    for (i = 0; i < n; i++) out[i] = t[n - 1 - i];
    out[n] = 0;
}

static void wait_before_retry(int attempt)
{
    int s = attempt < 10 ? attempt : 10;
    Sleep((DWORD)s * 1000);
}

/* Downloads one address. Returns 0, 3, 4 or 8, like wget. */
static int fetch(HINTERNET net, const char *url)
{
    int attempt, rc = 4;
    char num[24], num2[24];
    filename(url, g_name, sizeof g_name);
    request_url(url, g_req, sizeof g_req);
    if (g_outfile[0]) {
        lstrcpynA(g_path, g_outfile, sizeof g_path);
    } else if (g_folder[0]) {
        lstrcpyA(g_path, g_folder);
        if (g_path[lstrlenA(g_path) - 1] != '\\') lstrcatA(g_path, "\\");
        lstrcatA(g_path, g_name);
    } else {
        lstrcpyA(g_path, g_name);
    }
    if (g_noclobber && sizeof_file(g_path) >= 0) {
        wsprintfA(g_line, "File \"%s\" already there; not retrieving.", g_path);
        logline(g_line);
        return 0;
    }
    for (attempt = 1; attempt <= g_tries; attempt++) {
        HINTERNET h;
        HANDLE f;
        DWORD code = 0, len = sizeof code, flags, got, wrote;
        long long have, offset = 0, total = 0, expect = -1, flushed = 0;
        int resume, ok;
        have = sizeof_file(g_path);
        resume = (have > 0) && (g_continue || attempt > 1);
        g_hdr[0] = 0;
        if (resume) {
            numtext(have, num);
            wsprintfA(g_hdr, "Range: bytes=%s-\r\n", num);
        }
        flags = INTERNET_FLAG_RELOAD | INTERNET_FLAG_NO_CACHE_WRITE | INTERNET_FLAG_PRAGMA_NOCACHE |
                INTERNET_FLAG_NO_UI | INTERNET_FLAG_NO_COOKIES | INTERNET_FLAG_KEEP_CONNECTION;
        h = InternetOpenUrlA(net, g_req, g_hdr[0] ? g_hdr : NULL, g_hdr[0] ? (DWORD)-1 : 0, flags, 0);
        if (!h) {
            wsprintfA(g_line, "%s: unable to connect (Windows error %lu). Try %d of %d.", url, GetLastError(), attempt, g_tries);
            logline(g_line);
            rc = 4;
            if (attempt < g_tries) wait_before_retry(attempt);
            continue;
        }
        if (!HttpQueryInfoA(h, HTTP_QUERY_STATUS_CODE | HTTP_QUERY_FLAG_NUMBER, &code, &len, NULL)) code = 0;
        if (code == 416 && resume) {
            /* asked for the bytes after the end: we have it all */
            InternetCloseHandle(h);
            wsprintfA(g_line, "%s: the file is already fully retrieved; nothing to do.", url);
            logline(g_line);
            return 0;
        }
        if (code >= 400) {
            char reason[96];
            DWORD rl = sizeof reason;
            if (!HttpQueryInfoA(h, HTTP_QUERY_STATUS_TEXT, reason, &rl, NULL)) lstrcpyA(reason, "Error");
            InternetCloseHandle(h);
            wsprintfA(g_line, "%s: ERROR %lu: %s.", url, code, reason);
            logline(g_line);
            if (code >= 500 && attempt < g_tries) { rc = 8; wait_before_retry(attempt); continue; }
            return 8;
        }
        if (code != 200 && code != 206) {
            InternetCloseHandle(h);
            wsprintfA(g_line, "%s: unexpected answer from the server (%lu).", url, code);
            logline(g_line);
            rc = 8;
            if (attempt < g_tries) wait_before_retry(attempt);
            continue;
        }
        {
            char cl[32];
            DWORD cll = sizeof cl;
            if (HttpQueryInfoA(h, HTTP_QUERY_CONTENT_LENGTH, cl, &cll, NULL)) {
                const char *c = cl;
                expect = 0;
                while (*c >= '0' && *c <= '9') expect = expect * 10 + (*c++ - '0');
            }
        }
        if (code == 206 && resume) {
            f = CreateFileA(g_path, GENERIC_WRITE, FILE_SHARE_READ, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
            if (f != INVALID_HANDLE_VALUE) {
                LONG hi = (LONG)(have >> 32);
                SetFilePointer(f, (LONG)(have & 0xFFFFFFFF), &hi, FILE_BEGIN);
                SetEndOfFile(f);
            }
            offset = have;
        } else {
            f = CreateFileA(g_path, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
            offset = 0;
        }
        if (f == INVALID_HANDLE_VALUE) {
            InternetCloseHandle(h);
            wsprintfA(g_line, "Cannot write to \"%s\" (Windows error %lu).", g_path, GetLastError());
            logline(g_line);
            return 3;
        }
        ok = 1;
        for (;;) {
            got = 0;
            if (!InternetReadFile(h, g_buf, CHUNK, &got)) {
                wsprintfA(g_line, "%s: the connection was lost (Windows error %lu). Try %d of %d.", url, GetLastError(), attempt, g_tries);
                logline(g_line);
                ok = 0;
                break;
            }
            if (got == 0) break;
            if (!WriteFile(f, g_buf, got, &wrote, NULL) || wrote != got) {
                CloseHandle(f);
                InternetCloseHandle(h);
                wsprintfA(g_line, "Cannot write to \"%s\" (Windows error %lu). The drive may be full.", g_path, GetLastError());
                logline(g_line);
                return 3;
            }
            total += got;
            /* Windows 98 only tells others the new size of a file when it is flushed:
               the LAN Center reads the size to show how far along the download is */
            if (g_is9x && total - flushed >= 4 * 1024 * 1024) {
                FlushFileBuffers(f);
                flushed = total;
            }
        }
        CloseHandle(f);
        InternetCloseHandle(h);
        if (ok && expect >= 0 && total < expect) {
            numtext(offset + total, num);
            wsprintfA(g_line, "%s: the connection closed at byte %s. Try %d of %d.", url, num, attempt, g_tries);
            logline(g_line);
            ok = 0;
        }
        if (!ok) {
            rc = 4;
            if (attempt < g_tries) wait_before_retry(attempt);
            continue;
        }
        numtext(offset + total, num);
        numtext(total, num2);
        wsprintfA(g_line, "%s [%s bytes, %s new] -> \"%s\"", url, num, num2, g_path);
        logline(g_line);
        return 0;
    }
    return rc;
}

static void usage(void)
{
    MessageBoxA(NULL,
        "palmsget " VERSION " - downloads files for Palms LAN Center.\n\n"
        "palmsget [-c] [-nc] [-t tries] [-T seconds] [-o logfile] [-P folder] [-O file] [-i urlfile] [URL...]\n\n"
        "Palms LAN Center runs it by itself; there is no need to run it.",
        "palmsget", MB_OK | MB_ICONINFORMATION);
}

void start(void)
{
    char word[MAX_PATH * 2];
    char *p, *space;
    int i, result = 0, bad = 0;
    HINTERNET net;
    DWORD ms;

    g_is9x = (GetVersion() & 0x80000000) != 0;
    lstrcpynA(g_cmd, GetCommandLineA(), sizeof g_cmd);
    p = nextword(g_cmd, word, sizeof word);           /* the program itself */
    space = g_list;
    while (p && (p = nextword(p, word, sizeof word)) != NULL) {
        if (same(word, "-c")) g_continue = 1;
        else if (same(word, "-nc")) g_noclobber = 1;
        else if (same(word, "-nv") || same(word, "-q") || same(word, "--no-proxy")) { }
        else if (same(word, "-t")) { if ((p = nextword(p, word, sizeof word)) != NULL) g_tries = number(word); }
        else if (same(word, "-T")) { if ((p = nextword(p, word, sizeof word)) != NULL) g_timeout = number(word); }
        else if (same(word, "-o")) { if ((p = nextword(p, word, sizeof word)) != NULL) lstrcpynA(g_logname, word, sizeof g_logname); }
        else if (same(word, "-P")) { if ((p = nextword(p, word, sizeof word)) != NULL) lstrcpynA(g_folder, word, sizeof g_folder); }
        else if (same(word, "-O")) { if ((p = nextword(p, word, sizeof word)) != NULL) lstrcpynA(g_outfile, word, sizeof g_outfile); }
        else if (same(word, "-i")) { if ((p = nextword(p, word, sizeof word)) != NULL) lstrcpynA(g_urlfile, word, sizeof g_urlfile); }
        else if (word[0] == '-') bad = 1;
        else {
            int n = lstrlenA(word) + 1;
            if (space + n < g_list + sizeof g_list) {
                lstrcpyA(space, word);
                addurl(space);
                space += n;
            }
        }
        if (!p) break;
    }
    if (g_tries < 1) g_tries = 1;
    if (g_timeout < 1) g_timeout = 900;
    if (g_logname[0])
        g_log = CreateFileA(g_logname, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (g_urlfile[0] && !readlist(g_urlfile, &space, (int)(g_list + sizeof g_list - space))) {
        wsprintfA(g_line, "Unable to read the list of addresses \"%s\".", g_urlfile);
        logline(g_line);
        ExitProcess(3);
    }
    if (bad || g_nurls == 0) {
        if (g_log == INVALID_HANDLE_VALUE) usage();
        else logline("Nothing to download, or an option palmsget doesn't know.");
        ExitProcess(2);
    }
    if (g_folder[0]) CreateDirectoryA(g_folder, NULL);
    net = InternetOpenA("PalmsGet/" VERSION, INTERNET_OPEN_TYPE_DIRECT, NULL, NULL, 0);
    if (!net) {
        wsprintfA(g_line, "Unable to start the Internet functions (Windows error %lu). Internet Explorer 4 or later is needed.", GetLastError());
        logline(g_line);
        ExitProcess(4);
    }
    ms = (DWORD)g_timeout * 1000;
    InternetSetOptionA(net, INTERNET_OPTION_CONNECT_TIMEOUT, &ms, sizeof ms);
    InternetSetOptionA(net, INTERNET_OPTION_SEND_TIMEOUT, &ms, sizeof ms);
    InternetSetOptionA(net, INTERNET_OPTION_RECEIVE_TIMEOUT, &ms, sizeof ms);
    for (i = 0; i < g_nurls; i++) {
        int rc = fetch(net, g_urls[i]);
        /* like wget: the lower error numbers (2 and up) count for more */
        if (rc != 0 && (result == 0 || rc < result)) result = rc;
    }
    InternetCloseHandle(net);
    if (g_log != INVALID_HANDLE_VALUE) CloseHandle(g_log);
    ExitProcess((UINT)result);
}
