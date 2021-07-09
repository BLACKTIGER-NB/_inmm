#include <Windows.h>
#include <WinGDI.h>
#include <ddraw.h>
#include <Shlwapi.h>

#define _INMM_LOG_OUTPUT _DEBUG
//#define _INMM_PERF_LOG

#if _INMM_LOG_OUTPUT
#include <stdio.h>
#ifdef _INMM_PERF_LOG
#include <MMSystem.h>
#endif
#endif

static HDC hDesktopDC = NULL;
static HPALETTE hPalette = NULL;

// フォントテーブル
const int nMaxFontTable = 32;
typedef struct
{
	HFONT hFont;
	char szFaceName[LF_FACESIZE];
	int nHeight;
	int nBold;
	int nAdjustX;
	int nAdjustY;
} FONT_TABLE;
static FONT_TABLE fontTable[nMaxFontTable];

// プリセットカラーテーブル
const int nMaxColorTable = 'Z' - 'A' + 1;
typedef struct
{
	BYTE r;
	BYTE g;
	BYTE b;
} COLOR_TABLE;
static COLOR_TABLE colorTable[nMaxColorTable] =
{
	{   0,  64,   0 }, // A
	{   0,   0, 192 }, // B
	{ 128,   0,   0 }, // C
	{ 255, 255,   0 }, // D
	{ 255,   0, 255 }, // E
	{ 255, 128,   0 }, // F
	{   0, 100,   0 }, // G
	{ 128, 128, 128 }, // H
	{   0,   0, 255 }, // I
	{   0, 192,   0 }, // J
	{   0,   0,   0 }, // K
	{   0,   0,   0 }, // L
	{ 100, 255, 100 }, // M
	{ 255, 100, 100 }, // N
	{   0,   0,   0 }, // O
	{   0,   0,   0 }, // P
	{   0,   0,   0 }, // Q
	{ 240,   0,   0 }, // R
	{   0,   0,   0 }, // S
	{   0,   0,   0 }, // T
	{   0,   0,   0 }, // U
	{   0,   0,   0 }, // V
	{ 255, 255, 255 }, // W
	{   0,   0,   0 }, // X
	{ 208, 232, 248 }, // Y
	{   0,   0,   0 }, // Z
};

// コイン画像表示用のメモリデバイスコンテキストのハンドル
static HDC hDCCoin = NULL;

// コイン画像のビットマップハンドル
static HBITMAP hBitmapCoin = NULL;

// コイン画像読み込み前のビットマップハンドル退避先
static HBITMAP hBitmapCoinPrev = NULL;

// コイン画像のファイル名
static char szCoinFileName[MAX_PATH] = "Gfx\\Fonts\\Modern\\Additional\\coin.bmp";

// コイン画像の表示位置調整
static int nCoinAdjustX = 0;
static int nCoinAdjustY = 0;

// コイン画像の大きさ
static const int nCoinWidth = 12;
static const int nCoinHeight = 12;

// 最大表示可能行数
static int nMaxLines = 54;

// 単語内の最大表示可能文字数
#define NMAXWORDCHARS_NUMBER 24
static int nMaxWordChars =  NMAXWORDCHARS_NUMBER;

// WINMM.DLLのモジュールハンドル
static HMODULE hWinMmDll = NULL;

// 文字列処理バッファ
#define STRING_BUFFER_SIZE 1024

// ログ出力用ファイルポインタ
#if _INMM_LOG_OUTPUT
static FILE *fpLog = NULL;
#endif

// 公開関数
int WINAPI GetTextWidth(LPCBYTE lpString, int nMagicCode, DWORD dwFlags);
void WINAPI TextOutDC0(int x, int y, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags);
void WINAPI TextOutDC1(LPRECT lpRect, int x, int y, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags);
void WINAPI TextOutDC2(LPRECT lpRect, int *px, int *py, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags);
int WINAPI CalcLineBreak(LPBYTE lpBuffer, LPCBYTE lpString);
int WINAPI strnlen0(LPCBYTE lpString, int nMax);

// 内部関数
static int GetTextWidthWord(LPCBYTE lpString, int nLen, int nMagicCode);
static void TextOutWord(LPCBYTE lpString, int nLen, LPPOINT lpPoint, LPRECT lpRect, HDC hDC, int nMagicCode, COLORREF color, DWORD dwFlags);
static void ShowCoinImage(LPPOINT lpPoint, HDC hDC);
static int CalcColorWordWrap(LPBYTE lpBuffer, LPCBYTE lpString);
static int CalcAlphanumericWordWrap(LPBYTE lpBuffer, LPCBYTE lpString);
static void Init();
static void SetFontDefault();
static void LoadIniFile();
static void InitFont();
static void InitPalette();
static void InitCoinImage();
static void Terminate();
static int ReadLine(char *p, char *buf, int len);
//
static int CharUtf8toUtf32(LPCBYTE lpString, LPDWORD stringword);
static int Utf8CharLength(LPCBYTE utf8headbyte);
static bool Utf8Laterbytecheck(LPCBYTE utf8laterbyte);

// winmm.dllのAPI転送用
typedef MCIERROR (WINAPI *LPFNMCISENDCOMMANDA)(MCIDEVICEID, UINT, DWORD, DWORD);
typedef MMRESULT (WINAPI *LPFNTIMEBEGINPERIOD)(UINT);
typedef MMRESULT (WINAPI *LPFNTIMEGETDEVCAPS)(LPTIMECAPS, UINT);
typedef DWORD (WINAPI *LPFNTIMEGETTIME)(VOID);
typedef MMRESULT (WINAPI *LPFNTIMEKILLEVENT)(UINT);
typedef MMRESULT (WINAPI *LPFNTIMESETEVENT)(UINT, UINT, LPTIMECALLBACK, DWORD, UINT);

static LPFNMCISENDCOMMANDA lpfnMciSendCommandA = NULL;
static LPFNTIMEBEGINPERIOD lpfnTimeBeginPeriod = NULL;
static LPFNTIMEGETDEVCAPS lpfnTimeGetDevCaps = NULL;
static LPFNTIMEGETTIME lpfnTimeGetTime = NULL;
static LPFNTIMEKILLEVENT lpfnTimeKillEvent = NULL;
static LPFNTIMESETEVENT lpfnTimeSetEvent = NULL;

MCIERROR WINAPI _mciSendCommandA(MCIDEVICEID IDDevice, UINT uMsg, DWORD fdwCommand, DWORD dwParam);
MMRESULT WINAPI _timeBeginPeriod(UINT uPeriod);
MMRESULT WINAPI _timeGetDevCaps(LPTIMECAPS ptc, UINT cbtc);
DWORD WINAPI _timeGetTime(VOID);
MMRESULT WINAPI _timeKillEvent(UINT uTimerID);
MMRESULT WINAPI _timeSetEvent(UINT uDelay, UINT uResolution, LPTIMECALLBACK lpTimeProc, DWORD dwUser, UINT fuEvent);

//
// DLLメイン関数
//
// パラメータ
//   hInstDLL: DLL モジュールのハンドル
//   fdwReason: 関数を呼び出す理由
//   lpvReserved: 予約済み
// 戻り値
//   処理が成功すればTRUEを返す
//   
BOOL WINAPI DllMain(HINSTANCE hInstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	switch (fdwReason)
	{
	case DLL_PROCESS_ATTACH:
		DisableThreadLibraryCalls(hInstDLL);
		Init();
		break;
	case DLL_PROCESS_DETACH:
		Terminate();
		break;
	default:
		break;
	}
	return TRUE;
}

//
// 文字列幅を取得する (オリジナル版互換)
//
// パラメータ
//   lpString		対象文字列
//   nMagicCode		フォント指定用マジックコード
//   dwFlags		フラグ(1で太文字/2で影付き)
// 戻り値
//   文字列幅
//
int WINAPI GetTextWidth(LPCBYTE lpString, int nMagicCode, DWORD dwFlags)
{
#if _INMM_LOG_OUTPUT
#ifdef _INMM_PERF_LOG
	LARGE_INTEGER nFreq;
	LARGE_INTEGER nBefore;
	LARGE_INTEGER nAfter;
	memset(&nFreq, 0, sizeof(nFreq));
	memset(&nBefore, 0, sizeof(nBefore));
	memset(&nAfter, 0, sizeof(nAfter));
	QueryPerformanceFrequency(&nFreq);
	QueryPerformanceCounter(&nBefore);
#endif
#endif

	LPCBYTE s = lpString;
	LPCBYTE p = s;
	int nWidth = 0;
	int nMaxWidth = 0;
	int nLines = 0;

	if (fontTable[nMagicCode].hFont == NULL)
	{
		nMagicCode = 0;
	}
	HFONT hFontOld = (HFONT)SelectObject(hDesktopDC, fontTable[nMagicCode].hFont);

	// 先頭の連続したスペースは1文字を残して省略して計算する
	if  (*s == 0x20)
	{
		nWidth += GetTextWidthWord(p, 1, nMagicCode);
	}
	while (*s == 0x20)
	{
		s++;
		p=s;
	}

	while (*s != '\0')
	{
		switch (*s)
		{
		case 0xC2:
			// §(0xC2A7)
			if (*(s + 1) == 0xA7)
			{
				// §の後に英大文字が続けばプリセットの色指定
				if ((*(s + 2) >= 'A') && (*(s + 2) <= 'Z'))
				{
					if (s > p)
					{
						nWidth += GetTextWidthWord(p, s - p, nMagicCode);
					}
					s += 3;
					p = s;
					continue;
				}
				// §§ならば元の色に戻す
				if ((*(s + 2) == 0xC2) && (*(s + 3) == 0xA7))
				{
					if (s > p)
					{
						nWidth += GetTextWidthWord(p, s - p, nMagicCode);
					}
					s += 4;
					p = s;
					continue;
				}
				// 次の3文字が0～7ならば直値の色指定
				if ((*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7') && (*(s + 4) >= '0') && (*(s + 4) <= '7'))
				{
					if (s > p)
					{
						nWidth += GetTextWidthWord(p, s - p, nMagicCode);
					}
					s += 5;
					p = s;
					continue;
				}
				s += 2;
				continue;
			}
			// UTF-8の2バイト目
			if ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF))
			{
				s += 2;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0xA7: // ｧ(shiftjis)
			// ｧの後に英大文字が続けばプリセットの色指定
			if ((*(s + 1) >= 'A') && (*(s + 1) <= 'Z'))
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				s += 2;
				p = s;
				continue;
			}
			// ｧｧならば元の色に戻す
			if (*(s + 1) == 0xA7)
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				s += 2;
				p = s;
				continue;
			}
			// 次の3文字が0～7ならば直値の色指定
			if ((*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7'))
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				s += 4;
				p = s;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0xEF: // ｧ(0xEFBDA7)
			if ((*(s + 1) == 0xBD) && (*(s + 2) == 0xA7))
			{
				// ｧの後に英大文字が続けばプリセットの色指定
				if ((*(s + 3) >= 'A') && (*(s + 3) <= 'Z'))
				{
					if (s > p)
					{
						nWidth += GetTextWidthWord(p, s - p, nMagicCode);
					}
					s += 4;
					p = s;
					continue;
				}
				// ｧｧならば元の色に戻す
				if ((*(s + 3) == 0xEF) && (*(s + 4) == 0xBD) && (*(s + 5) == 0xA7))
				{
					if (s > p)
					{
						nWidth += GetTextWidthWord(p, s - p, nMagicCode);
					}
					s += 6;
					p = s;
					continue;
				}
				// 次の3文字が0～7ならば直値の色指定
				if ((*(s + 3) >= '0') && (*(s + 3) <= '7') && (*(s + 4) >= '0') && (*(s + 4) <= '7') && (*(s + 5) >= '0') && (*(s + 5) <= '7'))
				{
					if (s > p)
					{
						nWidth += GetTextWidthWord(p, s - p, nMagicCode);
					}
					s += 6;
					p = s;
					continue;
				}
			}
			// UTF-8の2-3バイト目
			if (((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)) && ((*(s + 2) >= 0x80) && (*(s + 2) <= 0xBF)))
			{
				s += 3;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case '%':
			// 次の3文字が0～7ならば直値の色指定
			if ((*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7'))
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				s += 4;
				p = s;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case '$':
			// コイン画像があれば幅を加算する
			if (hBitmapCoin != NULL)
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				nWidth += nCoinWidth;
				s++;
				p = s;
				continue;
			}
			// コイン画像がなければ単体の文字として扱う
			s++;
			continue;

		case 0x5C: // エスケープ記号 '\\'(shiftjis) or バックスラッシュ
			// エスケープ記号の後に%/$が続けば単体の文字として扱う
			if (*(s + 1) == '%' || *(s + 1) == '$')
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				s++;
				nWidth += GetTextWidthWord(s, 1, nMagicCode);
				s++;
				p = s;
				continue;
			}
			// エスケープ記号の後に§が続けば単体の文字として扱う
			if ((*(s + 1) == 0xC2) && (*(s + 2) == 0xA7))
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				s++;
				nWidth += GetTextWidthWord(s, 2, nMagicCode);
				s += 2;
				p = s;
				continue;
			}
			// エスケープ記号の後にｧが続けば単体の文字として扱う
			if ((*(s + 1) == 0xEF) && (*(s + 2) == 0xBD) && (*(s + 3) == 0xA7))
			{
				if (s > p)
				{
					nWidth += GetTextWidthWord(p, s - p, nMagicCode);
				}
				s++;
				nWidth += GetTextWidthWord(s, 2, nMagicCode);
				s += 3;
				p = s;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0x0A:
		case 0x0D:
			// 改行コード
			if (s > p)
			{
				nWidth += GetTextWidthWord(p, s - p, nMagicCode);
			}
			s++;
			p = s;
			if (nWidth > nMaxWidth)
			{
				nMaxWidth = nWidth;
			}
			nWidth = 0;
			nLines++;
			// 最大表示可能行数に到達していなければ計算を継続する
			if (nLines < nMaxLines)
			{
				continue;
			}
			// 最大表示可能行数に到達したら計算を打ち切る
			*(char *)(s - 1) = '\0';
			break;

		case 0x01:
		case 0x02:
		case 0x03:
		case 0x04:
		case 0x05:
		case 0x06:
		case 0x07:
		case 0x08:
		case 0x09:
		case 0x0B:
		case 0x0C:
		case 0x0E:
		case 0x0F:
		case 0x10:
		case 0x11:
		case 0x12:
		case 0x13:
		case 0x14:
		case 0x15:
		case 0x16:
		case 0x17:
		case 0x18:
		case 0x19:
		case 0x1A:
		case 0x1B:
		case 0x1C:
		case 0x1D:
		case 0x1E:
		case 0x1F:
			// 改行以外の制御コードならば読み飛ばす
			if (s > p)
			{
				nWidth += GetTextWidthWord(p, s - p, nMagicCode);
			}
			s++;
			p = s;
			continue;

		case 0xC0:
		case 0xC1:
		case 0xC3:
		case 0xC4:
		case 0xC5:
		case 0xC6:
		case 0xC7:
		case 0xC8:
		case 0xC9:
		case 0xCA:
		case 0xCB:
		case 0xCC:
		case 0xCD:
		case 0xCE:
		case 0xCF:
		case 0xD0:
		case 0xD1:
		case 0xD2:
		case 0xD3:
		case 0xD4:
		case 0xD5:
		case 0xD6:
		case 0xD7:
		case 0xD8:
		case 0xD9:
		case 0xDA:
		case 0xDB:
		case 0xDC:
		case 0xDD:
		case 0xDE:
		case 0xDF:
			// UTF-8の2バイト文字 (§がある0xC2開始以外)
			if ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF))
			{
				s += 2;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0xE0:
		case 0xE1:
		case 0xE2:
		case 0xE3:
		case 0xE4:
		case 0xE5:
		case 0xE6:
		case 0xE7:
		case 0xE8:
		case 0xE9:
		case 0xEA:
		case 0xEB:
		case 0xEC:
		case 0xED:
		case 0xEE:
			// UTF-8の3バイト文字 (ｧがある0xEF開始以外)
			if (((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)) && ((*(s + 2) >= 0x80) && (*(s + 2) <= 0xBF)))
			{
				s += 3;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0xF0:
		case 0xF1:
		case 0xF2:
		case 0xF3:
		case 0xF4:
		case 0xF5:
		case 0xF6:
		case 0xF7:
			// UTF-8の4バイト文字
			if (((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)) && ((*(s + 2) >= 0x80) && (*(s + 2) <= 0xBF)) && ((*(s + 3) >= 0x80) && (*(s + 3) <= 0xBF)))
			{
				s += 4;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		default:
			// それ以外ならば単体の文字として扱う
			s++;
			continue;
		}
		// 最大表示可能行数に到達した時に計算を打ち切るためのbreak
		// 続ける場合はswitch文の中にcontinueを書いてループの先頭へ飛ぶこと
		break;
	}
	if (s > p)
	{
		nWidth += GetTextWidthWord(p, s - p, nMagicCode);
	}
	if (nWidth > nMaxWidth)
	{
		nMaxWidth = nWidth;
	}
	if (nMaxWidth > 0)
	{
		nMaxWidth += ((dwFlags & 2) != 0) ? 1 : 0;
	}

	SelectObject(hDesktopDC, hFontOld);

#if _INMM_LOG_OUTPUT
#ifdef _INMM_PERF_LOG
	QueryPerformanceCounter(&nAfter);
	fprintf(fpLog, "[PERF] GetTextWidth: %lf\n", (double)(nAfter.QuadPart - nBefore.QuadPart) * 1000 / nFreq.QuadPart);
#endif
	fprintf(fpLog, "GetTextWidth: %s (%d) %X => %d\n", (char *)lpString, nMagicCode, dwFlags, nMaxWidth);
#endif

	return nMaxWidth;
}

//
// 単語の文字列幅を取得する
//
// パラメータ
//   lpString		対象文字列
//   nLen           文字列長(Byte単位)
//   nMagicCode		フォント指定用マジックコード
// 戻り値
//   文字列幅
//
int GetTextWidthWord(LPCBYTE lpString, int nLen, int nMagicCode)
{
	//文字列をUTF-8からUTF-16に変換する
	LPCBYTE s = lpString;
	DWORD u32chr;
	WORD p[256];
	int u8charlength = 0;
	int u8wordlength = 0;
	int i = 0;

	while (u8wordlength < nLen)
	{
		u8charlength = CharUtf8toUtf32(s, &u32chr);
		
		//不明文字なら処理を中断する
		if (u32chr == 0)
		{
			if (i != 0)
			{
				break;
			}
			else
			{
				return 0;
			}
		}

		u8wordlength += u8charlength;
		s += u8charlength;

		if (u32chr < 0x10000) {
			p[i++] = WORD(u32chr);
		}
		else
		{
			p[i++] = WORD(((u32chr - 0x10000) >> 10) + 0xD800);
			p[i++] = WORD(((u32chr - 0x10000) & 0x3FF) + 0xDC00);
		}
	}

	SIZE size;
	GetTextExtentPoint32W(hDesktopDC, (LPCWSTR)p, i, &size);

#if _INMM_LOG_OUTPUT
#ifndef _INMM_PERF_LOG
	char buf[STRING_BUFFER_SIZE];
	lstrcpyn(buf, (LPCSTR)lpString, STRING_BUFFER_SIZE);
	buf[nLen] = '\0';
	fprintf(fpLog, "  GetTextWidthWord %s %d %d => %d\n", buf, nLen, nMagicCode, size.cx);
#endif
#endif

	return size.cx;
}

//
// 文字列を出力する (オリジナル版互換)
//
// パラメータ
//   x				出力先X座標
//   y				出力先Y座標
//   lpString		対象文字列
//   lpDDS			DirectDrawサーフェースへのポインタ
//   nMagicCode		フォント指定用マジックコード
//   dwColor		色指定(R-G-Bを5bit-6bit-5bitのWORD値で)
//   dwFlags		フラグ(1で太文字/2で影付き)
//
void WINAPI TextOutDC0(int x, int y, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags)
{
	TextOutDC2(NULL, &x, &y, lpString, lpDDS, nMagicCode, dwColor, dwFlags);
}

//
// 文字列を出力する (クリッピング対応版)
//
// パラメータ
//   lpRect         クリッピング矩形(NULLの場合はクリッピングなし)
//   x				出力先X座標
//   y				出力先Y座標
//   lpString		対象文字列
//   lpDDS			DirectDrawサーフェースへのポインタ
//   nMagicCode		フォント指定用マジックコード
//   dwColor		色指定(R-G-Bを5bit-6bit-5bitのWORD値で)
//   dwFlags		フラグ(1で太文字/2で影付き)
//

void WINAPI TextOutDC1(LPRECT lpRect, int x, int y, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags)
{
	TextOutDC2(lpRect, &x, &y, lpString, lpDDS, nMagicCode, dwColor, dwFlags);
}

//
// 文字列を出力する (出力位置更新対応版)
//
// パラメータ
//   lpRect         クリッピング矩形(NULLの場合はクリッピングなし)
//   px				出力先X座標への参照
//   py				出力先Y座標への参照
//   lpString		対象文字列
//   lpDDS			DirectDrawサーフェースへのポインタ
//   nMagicCode		フォント指定用マジックコード
//   dwColor		色指定(R-G-Bを5bit-6bit-5bitのWORD値で)
//   dwFlags		フラグ(1で太文字/2で影付き)
//
void WINAPI TextOutDC2(LPRECT lpRect, int *px, int *py, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags)
{
#if _INMM_LOG_OUTPUT
	fprintf(fpLog, "TextOutDC2 %s (%d,%d) %X %d %X %X\n", (char *)lpString, *px, *py, (DWORD)lpDDS, nMagicCode, dwColor, dwFlags);
#ifdef _INMM_PERF_LOG
	LARGE_INTEGER nFreq;
	LARGE_INTEGER nBefore;
	LARGE_INTEGER nAfter;
	memset(&nFreq, 0, sizeof(nFreq));
	memset(&nBefore, 0, sizeof(nBefore));
	memset(&nAfter, 0, sizeof(nAfter));
	QueryPerformanceFrequency(&nFreq);
	QueryPerformanceCounter(&nBefore);
#endif
#endif

	LPCBYTE s = lpString;
	LPCBYTE p = s;
	COLORREF color = RGB(((dwColor & 0x0000F800) >> 11) << 3, ((dwColor & 0x000007E0) >> 5) << 2, (dwColor & 0x0000001F) << 3);
	COLORREF colorOld = color;

	HDC hDC;
	lpDDS->GetDC(&hDC);
	SetBkMode(hDC, TRANSPARENT);
	if (fontTable[nMagicCode].hFont == NULL)
	{
		nMagicCode = 0;
	}
	HFONT hFontOld = (HFONT)SelectObject(hDC, fontTable[nMagicCode].hFont);

	POINT pt = { *px, *py };

	// 先頭のスペースは省略する
	while (*s == 0x20)
	{
		s++;
		p = s;
	}

	while (*s != '\0')
	{
		switch (*s)
		{
		case 0xC2:
			// §(0xC2A7)
			if (*(s + 1) == 0xA7)
			{
				// §の後に英大文字が続けばプリセットの色指定
				if ((*(s + 2) >= 'A') && (*(s + 2) <= 'Z'))
				{
					// 出力する文字列があれば表示する
					if (s > p)
					{
						TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
					}
					colorOld = color;
					color = RGB(colorTable[*(s + 2) - 'A'].r, colorTable[*(s + 2) - 'A'].g, colorTable[*(s + 2) - 'A'].b);
					s += 3;
					p = s;
					continue;
				}
				// §§ならば元の色に戻す
				if ((*(s + 2) == 0xC2) && (*(s + 3) == 0xA7))
				{
					// 出力する文字列があれば表示する
					if (s > p)
					{
						TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
					}
					COLORREF temp = color;
					color = colorOld;
					colorOld = temp;
					s += 4;
					p = s;
					continue;
				}
				// 次の3文字が0～7ならば直値の色指定
				if ((*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7') && (*(s + 4) >= '0') && (*(s + 4) <= '7'))
				{
					// 出力する文字列があれば表示する
					if (s > p)
					{
						TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
					}
					colorOld = color;
					color = RGB((*(s + 4) - '0') << 5, (*(s + 3) - '0') << 5, (*(s + 2) - '0') << 5);
					s += 5;
					p = s;
					continue;
				}
				s += 2;
				continue;
			}
			// UTF-8の2バイト目
			if ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF))
			{
				s += 2;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0xA7: // ｧ(shiftjis)
			// ｧの後に英大文字が続けばプリセットの色指定
			if ((*(s + 1) >= 'A') && (*(s + 1) <= 'Z'))
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				colorOld = color;
				color = RGB(colorTable[*(s + 1) - 'A'].r, colorTable[*(s + 1) - 'A'].g, colorTable[*(s + 1) - 'A'].b);
				s += 2;
				p = s;
				continue;
			}
			// ｧｧならば元の色に戻す
			if (*(s + 1) == 0xA7)
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				COLORREF temp = color;
				color = colorOld;
				colorOld = temp;
				s += 2;
				p = s;
				continue;
			}
			// 次の3文字が0～7ならば直値の色指定
			if ((*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7'))
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				colorOld = color;
				color = RGB((*(s + 3) - '0') << 5, (*(s + 2) - '0') << 5, (*(s + 1) - '0') << 5);
				s += 4;
				p = s;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0xEF: // ｧ(0xEFBDA7)
			if ((*(s + 1) == 0xBD) && (*(s + 2) == 0xA7))
			{
				// ｧの後に英大文字が続けばプリセットの色指定
				if ((*(s + 3) >= 'A') && (*(s + 3) <= 'Z'))
				{
					// 出力する文字列があれば表示する
					if (s > p)
					{
						TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
					}
					colorOld = color;
					color = RGB(colorTable[*(s + 3) - 'A'].r, colorTable[*(s + 3) - 'A'].g, colorTable[*(s + 3) - 'A'].b);
					s += 4;
					p = s;
					continue;
				}
				// ｧｧならば元の色に戻す
				if ((*(s + 3) == 0xEF) && (*(s + 4) == 0xBD) && (*(s + 5) == 0xA7))
				{
					// 出力する文字列があれば表示する
					if (s > p)
					{
						TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
					}
					COLORREF temp = color;
					color = colorOld;
					colorOld = temp;
					s += 6;
					p = s;
					continue;
				}
				// 次の3文字が0～7ならば直値の色指定
				if ((*(s + 3) >= '0') && (*(s + 3) <= '7') && (*(s + 4) >= '0') && (*(s + 4) <= '7') && (*(s + 5) >= '0') && (*(s + 5) <= '7'))
				{
					// 出力する文字列があれば表示する
					if (s > p)
					{
						TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
					}
					colorOld = color;
					color = RGB((*(s + 3) - '0') << 5, (*(s + 2) - '0') << 5, (*(s + 1) - '0') << 5);
					s += 6;
					p = s;
					continue;
				}
			}
			// UTF-8の2-3バイト目
			if (((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)) && ((*(s + 2) >= 0x80) && (*(s + 2) <= 0xBF)))
			{
				s += 3;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case '%':
			// 次の3文字が0～7ならば直値の色指定
			if ((*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7'))
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				colorOld = color;
				color = RGB((*(s + 3) - '0') << 5, (*(s + 2) - '0') << 5, (*(s + 1) - '0') << 5);
				s += 4;
				p = s;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case '$':
			// コイン画像があれば表示する
			if (hBitmapCoin != NULL)
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				// コイン画像を表示する
				ShowCoinImage(&pt, hDC);
				s++;
				p = s;
				continue;
			}
			// コイン画像がなければ単体の文字として扱う
			s++;
			continue;

		case 0x5C: // エスケープ記号 '\\'(shiftjis) or バックスラッシュ
			// エスケープ記号の後に%/$が続けば単体の文字として扱う
			if (*(s + 1) == '%' || *(s + 1) == '$')
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				s++;
				TextOutWord(s, 1, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				s++;
				p = s;
				continue;
			}
			// エスケープ記号の後に§が続けば単体の文字として扱う
			if ((*(s + 1) == 0xA2) && (*(s + 2) == 0xA7))
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				s++;
				TextOutWord(s, 2, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				s += 2;
				p = s;
				continue;
			}
			// エスケープ記号の後にｧが続けば単体の文字として扱う
			if ((*(s + 1) == 0xEF) && (*(s + 2) == 0xBD) && (*(s + 3) == 0xA7))
			{
				// 出力する文字列があれば表示する
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				s++;
				TextOutWord(s, 2, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				s += 3;
				p = s;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0x01:
		case 0x02:
		case 0x03:
		case 0x04:
		case 0x05:
		case 0x06:
		case 0x07:
		case 0x08:
		case 0x09:
		case 0x0B:
		case 0x0C:
		case 0x0E:
		case 0x0F:
		case 0x10:
		case 0x11:
		case 0x12:
		case 0x13:
		case 0x14:
		case 0x15:
		case 0x16:
		case 0x17:
		case 0x18:
		case 0x19:
		case 0x1A:
		case 0x1B:
		case 0x1C:
		case 0x1D:
		case 0x1E:
		case 0x1F:
			// 出力する文字列があれば表示する
			if (s > p)
			{
				TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
			}
			// 制御コードを読み飛ばす
			s++;
			p = s;
			continue;

		case 0xC0:
		case 0xC1:
		case 0xC3:
		case 0xC4:
		case 0xC5:
		case 0xC6:
		case 0xC7:
		case 0xC8:
		case 0xC9:
		case 0xCA:
		case 0xCB:
		case 0xCC:
		case 0xCD:
		case 0xCE:
		case 0xCF:
		case 0xD0:
		case 0xD1:
		case 0xD2:
		case 0xD3:
		case 0xD4:
		case 0xD5:
		case 0xD6:
		case 0xD7:
		case 0xD8:
		case 0xD9:
		case 0xDA:
		case 0xDB:
		case 0xDC:
		case 0xDD:
		case 0xDE:
		case 0xDF:
			// UTF-8の2バイト文字 (§がある0xC2開始以外)
			if ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF))
			{
				s += 2;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0xE0:
		case 0xE1:
		case 0xE2:
		case 0xE3:
		case 0xE4:
		case 0xE5:
		case 0xE6:
		case 0xE7:
		case 0xE8:
		case 0xE9:
		case 0xEA:
		case 0xEB:
		case 0xEC:
		case 0xED:
		case 0xEE:
			// UTF-8の3バイト文字 (ｧがある0xEF開始以外)
			if (((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)) && ((*(s + 2) >= 0x80) && (*(s + 2) <= 0xBF)))
			{
				s += 3;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		case 0xF0:
		case 0xF1:
		case 0xF2:
		case 0xF3:
		case 0xF4:
		case 0xF5:
		case 0xF6:
		case 0xF7:
			// UTF-8の4バイト文字
			if (((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)) && ((*(s + 2) >= 0x80) && (*(s + 2) <= 0xBF)) && ((*(s + 3) >= 0x80) && (*(s + 3) <= 0xBF)))
			{
				s += 4;
				continue;
			}
			// それ以外ならば単体の文字として扱う
			s++;
			continue;

		default:
			// それ以外ならば単体の文字として扱う
			s++;
			continue;
		}
	}
	// 出力する文字列があれば表示する
	if (s > p)
	{
		TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
	}

	SelectObject(hDC, hFontOld);
	lpDDS->ReleaseDC(hDC);

	*px = pt.x;
	*py = pt.y;

#if _INMM_LOG_OUTPUT
#ifdef _INMM_PERF_LOG
	QueryPerformanceCounter(&nAfter);
	fprintf(fpLog, "[PERF] TextOutDC1: %lf\n", (double)(nAfter.QuadPart - nBefore.QuadPart) * 1000 / nFreq.QuadPart);
#endif
#endif
}

//
// 単語を出力する(途中で色変更はなし)
//
// パラメータ
//   lpString		対象文字列
//   nLen           文字列長(Byte単位)
//   lpPoint        出力先座標
//   lpRect         クリッピング矩形(NULLの場合はクリッピングなし)
//   hDC			デバイスコンテキストのハンドル
//   color			色指定
//   dwFlags		フラグ(1で太文字/2で影付き)
// 
void TextOutWord(LPCBYTE lpString, int nLen, LPPOINT lpPoint, LPRECT lpRect, HDC hDC, int nMagicCode, COLORREF color, DWORD dwFlags)
{
#if _INMM_LOG_OUTPUT
#ifndef _INMM_PERF_LOG
	fprintf(fpLog, "  TextOutWord %s %d (%d,%d) (%d,%d)-(%d,%d) %X %d %X %X\n", (char *)lpString, nLen, lpPoint->x, lpPoint->y,
		(lpRect != NULL) ? lpRect->left : 0, (lpRect != NULL) ? lpRect->top : 0, (lpRect != NULL) ? lpRect->right : 0, (lpRect != NULL) ? lpRect->bottom : 0,
		(DWORD)hDC, nMagicCode, (DWORD)color, dwFlags);
#endif
#endif

	int x = lpPoint->x + fontTable[nMagicCode].nAdjustX;
	int y = lpPoint->y + fontTable[nMagicCode].nAdjustY;


	//文字列をUTF-8からUTF-16に変換する
	LPCBYTE s = lpString;
	DWORD u32chr;
	WORD p[256];
	int u8charlength = 0;
	int u8wordlength = 0;
	int i = 0;

	while (u8wordlength < nLen)
	{
		u8charlength = CharUtf8toUtf32(s, &u32chr);

		//不明文字なら処理を中断する
		if (u32chr == 0)
		{
			if (i != 0)
			{
				break;
			}
			else
			{
				return;
			}
		}

		u8wordlength += u8charlength;
		s += u8charlength;

		if (u32chr < 0x10000) {
			p[i++] = WORD(u32chr);
		}
		else
		{
			p[i++] = WORD(((u32chr - 0x10000) >> 10) + 0xD800);
			p[i++] = WORD(((u32chr - 0x10000) & 0x3FF) + 0xDC00);
		}
	}


	// 影を付ける場合は出力位置の右下に黒で出力する
	bool bShadowed = ((dwFlags & 2) != 0);
	if (bShadowed)
	{
		SetTextColor(hDC, RGB(0, 0, 0));
		if (lpRect != NULL)
		{
			ExtTextOutW(hDC, x + 1, y + 1, ETO_CLIPPED, lpRect, (LPCWSTR)p, i, NULL);
		}
		else
		{
			TextOutW(hDC, x + 1, y + 1, (LPCWSTR)p, i);
		}
	}

	// 指定色で出力する
	SetTextColor(hDC, color);
	if (lpRect != NULL)
	{
		ExtTextOutW(hDC, x, y, ETO_CLIPPED, lpRect, (LPCWSTR)p, i, NULL);
	}
	else
	{
		TextOutW(hDC, x, y, (LPCWSTR)p, i);
	}
	
	// 出力位置を更新する
	SIZE size;
	GetTextExtentPoint32W(hDC, (LPCWSTR)p, i, &size);
	lpPoint->x += size.cx;
}


//
// 文字をUTF-32からUTF-16に変換する
//
// パラメータ
//   lpstringdw		変換前文字列
//   WordString		変換後文字列
//
void CharUtf32toUtf16(LPDWORD lpStringdw, LPWORD stringword)
{
}
//
// 文字をUTF-8からUTF32に変換する
//
// パラメータ
//   lpString			対象文字列
//   stringword			変換後後文字
//
int CharUtf8toUtf32(LPCBYTE lpString, LPDWORD stringdword)
{
	LPCBYTE s = lpString;
	LPDWORD p = stringdword;

	int charlength = Utf8CharLength(s);

	switch (charlength)
	{
	case 1:
		switch (*s)
		{
		case 0x80:
			// #EURO SIGN
			*p = 0x20AC;
			break;

		case 0x82:
			// #SINGLE LOW-9 QUOTATION MARK
			*p = 0x201A;
			break;

		case 0x83:
			// #LATIN SMALL LETTER F WITH HOOK
			*p = 0x0192;
			break;

		case 0x84:
			// #DOUBLE LOW-9 QUOTATION MARK
			*p = 0x201E;
			break;

		case 0x85:
			// #HORIZONTAL ELLIPSIS
			*p = 0x2026;
			break;

		case 0x86:
			// #DAGGER
			*p = 0x2020;
			break;

		case 0x87:
			// #DOUBLE DAGGER
			*p = 0x2021;
			break;

		case 0x88:
			// #MODIFIER LETTER CIRCUMFLEX ACCENT
			*p = 0x02C6;
			break;

		case 0x89:
			// #PER MILLE SIGN
			*p = 0x2030;
			break;

		case 0x8A:
			// #LATIN CAPITAL LETTER S WITH CARON
			*p = 0x0160;
			break;

		case 0x8B:
			// #SINGLE LEFT-POINTING ANGLE QUOTATION MARK
			*p = 0x2039;
			break;

		case 0x8C:
			// #LATIN CAPITAL LIGATURE OE
			*p = 0x0152;
			break;

		case 0x8E:
			// #LATIN CAPITAL LETTER Z WITH CARON
			*p = 0x017D;
			break;

		case 0x91:
			// #LEFT SINGLE QUOTATION MARK
			*p = 0x2018;
			break;

		case 0x92:
			// #RIGHT SINGLE QUOTATION MARK
			*p = 0x2019;
			break;

		case 0x93:
			// #LEFT DOUBLE QUOTATION MARK
			*p = 0x201C;
			break;

		case 0x94:
			// #RIGHT DOUBLE QUOTATION MARK
			*p = 0x201D;
			break;

		case 0x95:
			// #BULLET
			*p = 0x2022;
			break;

		case 0x96:
			// #EN DASH
			*p = 0x2013;
			break;

		case 0x97:
			// #EM DASH
			*p = 0x2014;
			break;

		case 0x98:
			// #SMALL TILDE
			*p = 0x02DC;
			break;

		case 0x99:
			// #TRADE MARK SIGN
			*p = 0x2122;
			break;

		case 0x9A:
			// #LATIN SMALL LETTER S WITH CARON
			*p = 0x0161;
			break;

		case 0x9B:
			// #SINGLE RIGHT-POINTING ANGLE QUOTATION MARK
			*p = 0x203A;
			break;

		case 0x9C:
			// #LATIN SMALL LIGATURE OE
			*p = 0x0153;
			break;

		case 0x9E:
			// #LATIN SMALL LETTER Z WITH CARON
			*p = 0x017E;
			break;

		case 0x9F:
			// #LATIN CAPITAL LETTER Y WITH DIAERESIS
			*p = 0x0178;
			break;

		default:
			*p = *s;
		}
		return charlength;
		break;

	case 2:
		// UTF-8の2バイト文字
		if ((*(s+1) >= 0x80) && (*(s+1) <= 0xBF))
		{
			*p = (*s & 0x1F) << 6;
			*p |= (*(s+1) & 0x3F);
			return charlength;
			break;
		}
		// 不明な文字は1バイト文字として処理する
		*p = *s;
		return 1;
		break;

	case 3:
		if (((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)) && ((*(s + 2) >= 0x80) && (*(s + 2) <= 0xBF)))
		{
			*p = (*s & 0x0F) << 12;
			*p |= (*(s+1) & 0x3F) << 6;
			*p |= (*(s+2) & 0x3F);
			return charlength;
			break;
		}
		// 不明な文字は1バイト文字として処理する
		*p = *s;
		return 1;
		break;

	case 4:
		if (((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)) && ((*(s + 2) >= 0x80) && (*(s + 2) <= 0xBF)) && ((*(s + 3) >= 0x80) && (*(s + 3) <= 0xBF)))
		{
			*p = (*s & 0x07) << 18;
			*p |= (*(s+1) & 0x3F) << 12;
			*p |= (*(s+2) & 0x3F) << 6;
			*p |= (*(s+3) & 0x3F);
			return charlength;
			break;
		}
		// 不明な文字は1バイト文字として処理する
		*p = *s;
		return 1;
		break;

	default:
		//不明な文字は1バイト文字として処理する
		*p = *s;
		return 1;
		break;
	}
}

//
//utf-8の先頭バイトから文字数を返す
//
// パラメータ
//   utf8headbyte		対象バイト
//
int Utf8CharLength(LPCBYTE utf8headbyte)
{
	//1バイトutf-8は0x00から0x7Fだが 0x80-0xC1も通すようにするかもしれない
	//暫定的にwindows-1952の0x80-0xBFも通すようにしてみる
	//暫定的に0xC0-0xC1も通してみる
	if ((*utf8headbyte >= 0x00) && (*utf8headbyte <= 0xC1))
	{
		return 1;
	}
	if ((*utf8headbyte >= 0xC2) && (*utf8headbyte <= 0xDF))
	{
		return 2;
	}
	if ((*utf8headbyte >= 0xE0) && (*utf8headbyte <= 0xEF))
	{
		return 3;
	}
	if ((*utf8headbyte >= 0xF0) && (*utf8headbyte <= 0xF7))
	{
		return 4;
	}
	return 0;
}

//
//utf-8の2バイト目以降か否かの判定
//
// パラメータ
//   utf8laterbyte		対象バイト
//
bool Utf8Laterbytecheck(LPCBYTE utf8laterbyte)
{
	return  ((*utf8laterbyte >= 0x80) && (*utf8laterbyte <= 0x7BF));
}


//
// コイン画像を表示する
//
// パラメータ
//   lpPoint        出力先座標
//   hDC			デバイスコンテキストのハンドル
//
void ShowCoinImage(LPPOINT lpPoint, HDC hDC)
{
#if _INMM_LOG_OUTPUT
#ifndef _INMM_PERF_LOG
	fprintf(fpLog, "  ShowCoinImage (%d,%d)\n", lpPoint->x, lpPoint->y);
#endif
#endif

	COLORREF crTransparent = GetPixel(hDCCoin, 0, 0);
	TransparentBlt(hDC, lpPoint->x + nCoinAdjustX, lpPoint->y + nCoinAdjustY, nCoinWidth, nCoinHeight, hDCCoin, 0, 0, nCoinWidth, nCoinHeight, (UINT)crTransparent);
	lpPoint->x += nCoinWidth;
}

//
// 改行位置を計算する (禁則処理対応版)
//
// パラメータ
//   lpBuffer		文字列処理バッファ
//   lpString		対象文字列
// 戻り値
//   処理したバイト数
//
int WINAPI CalcLineBreak(LPBYTE lpBuffer, LPCBYTE lpString)
{
#if _INMM_LOG_OUTPUT
#ifdef _INMM_PERF_LOG
	LARGE_INTEGER nFreq;
	LARGE_INTEGER nBefore;
	LARGE_INTEGER nAfter;
	memset(&nFreq, 0, sizeof(nFreq));
	memset(&nBefore, 0, sizeof(nBefore));
	memset(&nAfter, 0, sizeof(nAfter));
	QueryPerformanceFrequency(&nFreq);
	QueryPerformanceCounter(&nBefore);
#endif
#endif
	LPBYTE p = lpBuffer;
	LPCBYTE s = lpString;
	int len;

	while (*s != '\0')
	{
		// 行末禁則文字/制御文字の処理
		switch (*s)
		{
		case '(':
		case '<':
		case '[':
		case '`':
		case '{':
			// 行末禁則文字ならば次の文字も一緒に処理する
			*p++ = *s++;
			continue;

		case 0xC2:
			*p++ = *s++;
			switch (*s)
			{
			case 0xA7: // §(0xC2A7)
				*p++ = *s++;
				// §の後に英大文字が続けばプリセットの色指定
				if ((*s >= 'A') && (*s <= 'Z'))
				{
					// §Wは白色なので分離禁則対象外
					if (*s == 'W')
					{
						*p++ = *s++;
						// 色指定の次の文字も一緒に処理する
						continue;
					}
					// 色指定後の分離禁則処理
					*p++ = *s++;
					len = CalcColorWordWrap(p, s);
					p += len;
					s += len;
					break;
				}
				// §§ならば元の色に戻す
				if ((*s == 0xC2) && (*(s + 1) == 0xA7))
				{
					*p++ = *s++;
					*p++ = *s++;
					// 色指定の次の文字も一緒に処理する
					continue;
				}
				// 次の3文字が0～7ならば直値の色指定
				if ((*s >= '0') && (*s <= '7') && (*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7'))
				{
					*p++ = *s++;
					*p++ = *s++;
					*p++ = *s++;
					// 色指定の次の文字も一緒に処理する
					continue;
				}
				break;

			case 0xB1: // ±(0xC2B1)
				*p++ = *s++;
				// 連英数字の分離禁則処理
				len = CalcAlphanumericWordWrap(p, s);
				p += len;
				s += len;
				break;

			default:
				// UTF-8の2バイト目
				if ((*s >= 0x80) && (*s <= 0xBF))
				{
					*p++ = *s++;
					break;
				}
				// 不明な文字は1バイト文字として処理する
				break;
			}
			break;

		case 0xE2:
			*p++ = *s++;
			switch (*s)
			{
			case 0x80:
				switch (*(s + 1))
				{
				case 0x98: // ‘
				case 0x9C: // “
					// 行末禁則文字ならば次の文字も一緒に処理する
					*p++ = *s++;
					*p++ = *s++;
					continue;
				default:
					// UTF-8の3バイト目
					if ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF))
					{
						*p++ = *s++;
						*p++ = *s++;
						break;
					}
					break;
				}
				// 不明な文字は1バイト文字として処理する
				break;

			default:
					// UTF-8の2-3バイト目
				if (((*s >= 0x80) && (*s <= 0xBF)) && ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)))
				{
					*p++ = *s++;
					*p++ = *s++;
					break;
				}
				// 不明な文字は1バイト文字として処理する
				break;
			}
			break;

		case 0xE3:
			*p++ = *s++;
			switch (*s)
			{
			case 0x80:
				switch (*(s + 1)) // “
				{
				case 0x94: // 〔
				case 0x88: // 〈
				case 0x8A: // 《
				case 0x8C: // 「
				case 0x8E: // 『
				case 0x90: // 【
				case 0x9D: // 〝
					// 行末禁則文字ならば次の文字も一緒に処理する
					*p++ = *s++;
					*p++ = *s++;
					continue;
				default:
					// UTF-8の3バイト目
					if ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF))
					{
						*p++ = *s++;
						*p++ = *s++;
						break;
					}
					break;
				}
				// 不明な文字は1バイト文字として処理する
				break;

			default:
					// UTF-8の2-3バイト目
				if (((*s >= 0x80) && (*s <= 0xBF)) && ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)))
				{
					*p++ = *s++;
					*p++ = *s++;
					break;
				}
				// 不明な文字は1バイト文字として処理する
				break;
			}
			break;

		case 0xEF:
			*p++ = *s++;
			switch (*s)
			{
			case 0xBC:
				switch (*(s + 1))
				{
				case 0x88: // （
				case 0xBB: // ［
				case 0x9C: // ＜
				case 0x84: // ＄
					// 行末禁則文字ならば次の文字も一緒に処理する
					*p++ = *s++;
					*p++ = *s++;
					continue;

				case 0x8B: // ＋
				case 0x8D: // －
					*p++ = *s++;
					*p++ = *s++;
					// 連英数字の分離禁則処理
					len = CalcAlphanumericWordWrap(p, s);
					p += len;
					s += len;
				break;

				case 0x90: // ０
				case 0x91: // １
				case 0x92: // ２
				case 0x93: // ３
				case 0x94: // ４
				case 0x95: // ５
				case 0x96: // ６
				case 0x97: // ７
				case 0x98: // ８
				case 0x99: // ９
					*p++ = *s++;
					*p++ = *s++;
					// 連英数字の分離禁則処理
					len = CalcAlphanumericWordWrap(p, s);
					p += len;
					s += len;
				break;

				default:
					// UTF-8の3バイト目
					if ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF))
					{
						*p++ = *s++;
						*p++ = *s++;
						break;
					}
					break;
				}
				// 不明な文字は1バイト文字として処理する
				break;

			case 0xBD:
				switch (*(s + 1))
				{
				case 0x80: // ｀
				case 0x9B: // ｛
				case 0xA2: // ｢
					// 行末禁則文字ならば次の文字も一緒に処理する
					*p++ = *s++;
					*p++ = *s++;
					continue;

				case 0xA7: // ｧ
					*p++ = *s++;
					*p++ = *s++;
					// ｧの後に英大文字が続けばプリセットの色指定
					if ((*s >= 'A') && (*s <= 'Z'))
					{
						// ｧWは白色なので分離禁則対象外
						if (*s == 'W')
						{
							*p++ = *s++;
							// 色指定の次の文字も一緒に処理する
							continue;
						}
						// 色指定後の分離禁則処理
						*p++ = *s++;
						len = CalcColorWordWrap(p, s);
						p += len;
						s += len;
						break;
					}
					// ｧｧならば元の色に戻す
					if ((*s == 0xEF) && (*(s + 1) == 0xBD) && (*(s + 2) == 0xA7))
					{
						*p++ = *s++;
						*p++ = *s++;
						*p++ = *s++;
						// 色指定の次の文字も一緒に処理する
						continue;
					}
					// 次の3文字が0～7ならば直値の色指定
					if ((*s >= '0') && (*s <= '7') && (*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7'))
					{
						*p++ = *s++;
						*p++ = *s++;
						*p++ = *s++;
						// 色指定の次の文字も一緒に処理する
						continue;
					}
					break;

				default:
					// UTF-8の3バイト目
					if ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF))
					{
						*p++ = *s++;
						*p++ = *s++;
						break;
					}
					break;
				}
				// 不明な文字は1バイト文字として処理する
				break;

			case 0xBF:
				switch (*(s + 1))
				{
				case 0xA5: // ￥
					// 行末禁則文字ならば次の文字も一緒に処理する
					*p++ = *s++;
					*p++ = *s++;
					continue;
				default:
					// UTF-8の3バイト目
					if ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF))
					{
						*p++ = *s++;
						*p++ = *s++;
						break;
					}
					break;
				}
				// 不明な文字は1バイト文字として処理する
				break;

			default:
					// UTF-8の2-3バイト目
				if (((*s >= 0x80) && (*s <= 0xBF)) && ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)))
				{
					*p++ = *s++;
					*p++ = *s++;
					break;
				}
				// 不明な文字は1バイト文字として処理する
				break;
			}
			break;

		case 0xA7: // ｧ
			*p++ = *s++;
			// ｧの後に英大文字が続けばプリセットの色指定
			if ((*s >= 'A') && (*s <= 'Z'))
			{
				// ｧWは白色なので分離禁則対象外
				if (*s == 'W')
				{
					*p++ = *s++;
					// 色指定の次の文字も一緒に処理する
					continue;
				}
				// 色指定後の分離禁則処理
				*p++ = *s++;
				len = CalcColorWordWrap(p, s);
				p += len;
				s += len;
				break;
			}
			// ｧｧならば元の色に戻す
			if (*s == 0xA7)
			{
				*p++ = *s++;
				// 色指定の次の文字も一緒に処理する
				continue;
			}
			// 次の3文字が0～7ならば直値の色指定
			if ((*s >= '0') && (*s <= '7') && (*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7'))
			{
				*p++ = *s++;
				*p++ = *s++;
				*p++ = *s++;
				// 色指定の次の文字も一緒に処理する
				continue;
			}
			break;

		case '%':
			*p++ = *s++;
			// 次の3文字が0～7ならば直値の色指定
			if ((*s >= '0') && (*s <= '7') && (*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7'))
			{
				*p++ = *s++;
				*p++ = *s++;
				*p++ = *s++;
				// 色指定の次の文字も一緒に処理する
				continue;
			}
			break;

		case 0x5C: // エスケープ記号 '\\'(shiftjis) or バックスラッシュ
			*p++ = *s++;
			switch (*s)
			{
			case '$':
				// エスケープ記号の後に$が続けば単体の行末禁則文字として扱う
				*p++ = *s++;
				continue;

			case '%':
				// エスケープ記号の後に%が続けば単体の文字として扱う
				*p++ = *s++;
				break;

			case 0xEF: // ｧ
				if ((*(s + 1) == 0xBD) && (*(s + 2) == 0xA7))
				{
					// エスケープ記号の後にｧが続けば単体の文字として扱う
					*p++ = *s++;
					*p++ = *s++;
					*p++ = *s++;
				}
				else
				{
					// 単体の\は行末禁則文字として扱う
					continue;
				}
				break;

			case 0xC2:
				if (*(s + 1) == 0xA7)
				{
					// エスケープ記号の後に§が続けば単体の文字として扱う
					*p++ = *s++;
					*p++ = *s++;
				}
				else
				{
					// 単体の\は行末禁則文字として扱う
					continue;
				}
				break;

			default:
				// 単体の\は行末禁則文字として扱う
				continue;
			}
			break;

		case '0':
		case '1':
		case '2':
		case '3':
		case '4':
		case '5':
		case '6':
		case '7':
		case '8':
		case '9':
		case '-':
		case '+':
		case 'A':
		case 'B':
		case 'C':
		case 'D':
		case 'E':
		case 'F':
		case 'G':
		case 'H':
		case 'I':
		case 'J':
		case 'K':
		case 'L':
		case 'M':
		case 'N':
		case 'O':
		case 'P':
		case 'Q':
		case 'R':
		case 'S':
		case 'T':
		case 'U':
		case 'V':
		case 'W':
		case 'X':
		case 'Y':
		case 'Z':
		case 'a':
		case 'b':
		case 'c':
		case 'd':
		case 'e':
		case 'f':
		case 'g':
		case 'h':
		case 'i':
		case 'j':
		case 'k':
		case 'l':
		case 'm':
		case 'n':
		case 'o':
		case 'p':
		case 'q':
		case 'r':
		case 's':
		case 't':
		case 'u':
		case 'v':
		case 'w':
		case 'x':
		case 'y':
		case 'z':
		case 0x27: // #APOSTROPHE
			*p++ = *s++;
			// 連英数字の分離禁則処理
			len = CalcAlphanumericWordWrap(p, s);
			p += len;
			s += len;
			break;

		case 0x0A:
		case 0x0D:
			// 禁則文字の次の改行コードは処理しない
			// これを入れないと色指定の直後の改行が処理されない
			break;

		case 0xC0:
		case 0xC1:
		case 0xC3:
		case 0xC4:
		case 0xC5:
		case 0xC6:
		case 0xC7:
		case 0xC8:
		case 0xC9:
		case 0xCA:
		case 0xCB:
		case 0xCC:
		case 0xCD:
		case 0xCE:
		case 0xCF:
		case 0xD0:
		case 0xD1:
		case 0xD2:
		case 0xD3:
		case 0xD4:
		case 0xD5:
		case 0xD6:
		case 0xD7:
		case 0xD8:
		case 0xD9:
		case 0xDA:
		case 0xDB:
		case 0xDC:
		case 0xDD:
		case 0xDE:
		case 0xDF:
			*p++ = *s++;
			// UTF-8の2バイト文字 (§がある0xC2開始以外)
			if ((*s >= 0x80) && (*s <= 0xBF))
			{
				*p++ = *s++;
				break;
			}
			// 不明な文字は1バイト文字として処理する
			break;

		case 0xE0:
		case 0xE1:
		case 0xE4:
		case 0xE5:
		case 0xE6:
		case 0xE7:
		case 0xE8:
		case 0xE9:
		case 0xEA:
		case 0xEB:
		case 0xEC:
		case 0xED:
		case 0xEE:
			*p++ = *s++;
			// UTF-8の3バイト文字 (禁則文字があるがある0xE2,xE3,0xEF開始以外)
			if (((*s >= 0x80) && (*s <= 0xBF)) && ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)))
			{
			*p++ = *s++;
			*p++ = *s++;
				break;
			}
			// 不明な文字は1バイト文字として処理する
			break;

		case 0xF0:
		case 0xF1:
		case 0xF2:
		case 0xF3:
		case 0xF4:
		case 0xF5:
		case 0xF6:
		case 0xF7:
			*p++ = *s++;
			// UTF-8の4バイト文字
			if (((*s >= 0x80) && (*s <= 0xBF)) && ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)) && ((*(s + 2) >= 0x80) && (*(s + 2) <= 0xBF)))
			{
			*p++ = *s++;
			*p++ = *s++;
			*p++ = *s++;
				break;
			}
			// 不明な文字は1バイト文字として処理する
			break;

		default:
			*p++ = *s++;
			break;
		}


		// 行頭禁則文字/分離禁則文字の処理
		switch (*s)
		{
		case '!':
		case '%':
		case '\'':
		case '-':
		case '.':
		case ',':
		case ')':
		case ':':
		case ';':
		case '=':
		case '>':
		case '?':
		case ']':
		case '}':
		case '~':
			// 次が行頭禁則文字ならば一緒に処理する
			continue;

		case '/':
			// 次が分離禁則文字ならばその次も含めて一緒に処理する
			*p++ = *s++;
			continue;

		case 0x5c: // エスケープ記号 '\\'(shiftjis) or バックスラッシュ
			// 次が行頭禁則文字ならば一緒に処理する
			if ((*(s + 1) == 0xEF) && (*(s + 2) == 0xBD) && (*(s + 3) == 0xA7)) // ｧ
			{
				continue;
			}
			break;

		case 0xC2:
			switch (*(s + 1))
			{
			case 0xB4: // ´
				// 次が行頭禁則文字ならば一緒に処理する
					continue;
			}
			break;

		case 0xE2:
			switch (*(s + 1))
			{
			case 0x80:
				switch (*(s + 2))
				{
				case 0x99: // ’
				case 0x9D: // ”
					// 次が行頭禁則文字ならば一緒に処理する
					continue;
				case 0x95: // ―
				case 0xA6: // …
				case 0xA5: // ‥
					// 次が分離禁則文字ならばその次も含めて一緒に処理する
					*p++ = *s++;
					*p++ = *s++;
					*p++ = *s++;
					continue;
				}
				break;
			case 0x84:
				switch (*(s + 2))
				{
				case 0xB3: // ℃
					// 次が行頭禁則文字ならば一緒に処理する
					continue;
				}
				break;
			}
			break;

		case 0xE3:
			switch (*(s + 1))
			{
			case 0x80:
				switch (*(s + 2))
				{
				case 0x81: // 、
				case 0x82: // 。
				case 0x83: // 〃
				case 0x85: // 々
				case 0x90: // ‐
				case 0x95: // 〕
				case 0x89: // 〉
				case 0x8B: // 》
				case 0x8D: // 」
				case 0x8F: // 』
				case 0x91: // 】
				case 0x9F: // 〟
					// 次が行頭禁則文字ならば一緒に処理する
					continue;
				}
				break;
			case 0x81:
				switch (*(s + 2))
				{
				case 0x81: // ぁ
				case 0x83: // ぃ
				case 0x85: // ぅ
				case 0x87: // ぇ
				case 0x89: // ぉ
				case 0xA3: // っ
					// 次が行頭禁則文字ならば一緒に処理する
					continue;
				}
				break;
			case 0x82:
				switch (*(s + 2))
				{
				case 0x9B: // ゛
				case 0x9C: // ゜
				case 0x9D: // ゝ
				case 0x9E: // ゞ
				case 0x83: // ゃ
				case 0x85: // ゅ
				case 0x87: // ょ
				case 0x8E: // ゎ
				case 0xA1: // ァ
				case 0xA3: // ィ
				case 0xA5: // ゥ
				case 0xA7: // ェ
				case 0xA9: // ォ
					// 次が行頭禁則文字ならば一緒に処理する
					continue;
				}
				break;
			case 0x83:
				switch (*(s + 2))
				{
				case 0xBB: // ・
				case 0xBD: // ヽ
				case 0xBE: // ヾ
				case 0xBC: // ー
				case 0x83: // ッ
				case 0xA3: // ャ
				case 0xA5: // ュ
				case 0xA7: // ョ
				case 0xAE: // ヮ
				case 0xB5: // ヵ
				case 0xB6: // ヶ
					// 次が行頭禁則文字ならば一緒に処理する
					continue;
				}
				break;
			}
			break;

		case 0xEF:
			switch (*(s + 1))
			{
			case 0xBC:
				switch (*(s + 2))
				{
				case 0x8C: // ，
				case 0x8E: // ．
				case 0x9A: // ：
				case 0x9B: // ；
				case 0x9F: // ？
				case 0x81: // ！
				case 0x89: // ）
				case 0xBD: // ］
				case 0x9D: // ＝
				case 0x9E: // ＞
				case 0x85: // ％
					// 次が行頭禁則文字ならば一緒に処理する
					continue;
				case 0x8F: // ／
					// 次が分離禁則文字ならばその次も含めて一緒に処理する
					*p++ = *s++;
					*p++ = *s++;
					*p++ = *s++;
					continue;
				}
				break;
			case 0xBD:
				switch (*(s + 2))
				{
				case 0xA1: // ｡
				case 0xA3: // ｣
				case 0xA4: // ､
				case 0xA5: // ･
				case 0xA8: // ｨ
				case 0xA9: // ｩ
				case 0xAA: // ｪ
				case 0xAB: // ｫ
				case 0xAC: // ｬ
				case 0xAD: // ｭ
				case 0xAE: // ｮ
				case 0xAF: // ｯ
				case 0xB0: // ｰ
				case 0x9E: // ～
				case 0x9D: // ｝ 
					// 次が行頭禁則文字ならば一緒に処理する
					continue;
				}
				break;
			case 0xBE:
				switch (*(s + 2))
				{
				case 0x9E: // ﾞ
				case 0x9F: // ﾟ
					// 次が行頭禁則文字ならば一緒に処理する
					continue;
				}
				break;
			case 0xBF:
				switch (*(s + 2))
				{
				case 0xA0: // ￠
					// 次が行頭禁則文字ならば一緒に処理する
					continue;
				}
				break;
			}
			break;
		}
		break;
	}
	*p = '\0';

#if _INMM_LOG_OUTPUT
#ifdef _INMM_PERF_LOG
	QueryPerformanceCounter(&nAfter);
	fprintf(fpLog, "[PERF] CalcLineBreak: %lf\n", (double)(nAfter.QuadPart - nBefore.QuadPart) * 1000 / nFreq.QuadPart);
#endif
	fprintf(fpLog, "CalcLineBreak: %s (%d) %s\n", (char *)lpBuffer, p - lpBuffer, (char *)lpString);
#endif

	return p - lpBuffer;
}

//
// 色指定後の分離禁則を計算する
//
// パラメータ
//   lpBuffer		文字列処理バッファ
//   lpString		対象文字列
// 戻り値
//   処理したバイト数
//
int CalcColorWordWrap(LPBYTE lpBuffer, LPCBYTE lpString)
{
	LPBYTE p = lpBuffer;
	LPCBYTE s = lpString;

	while (*s != '\0')
	{
		// 単語内の最大表示可能文字数を超える場合その場で区切る(保険)
		if (p - lpBuffer > nMaxWordChars)
		{
			break;
		}
		switch (*s)
		{
		case 0xA7: // ｧ
			// 色指定の直前で区切る
			if ((*(s + 1) >= 'A') && (*(s + 1) <= 'Z'))
			{
				break;
			}
			if (*(s + 1) == 0xA7)
			{
				break;
			}
			// 色指定でなければ通常の文字として扱う
			*p++ = *s++;
			continue;

		case 0xC2:
			if (*(s + 1) == 0xA7) // §
			{
				// 色指定の直前で区切る
				if ((*(s + 2) >= 'A') && (*(s + 2) <= 'Z'))
				{
					break;
				}
				if ((*(s + 2) == 0xC2) && (*(s + 3) == 0xA7))
				{
					break;
				}
				// 色指定でなければ通常の文字として扱う
				*p++ = *s++;
				*p++ = *s++;
				continue;
			}
			*p++ = *s++;
			// UTF-8の2バイト目
			if ((*s >= 0x81) && (*s <= 0xBF))
			{
				*p++ = *s++;
				continue;
			}
			// 不明な文字は1バイト文字として処理する
			continue;

		case 0xEF: // ｧ
			// 色指定の直前で区切る
			if ((*(s + 1) == 0xBD) && (*(s + 2) == 0xA7))
			{
				if ((*(s + 3) >= 'A') && (*(s + 3) <= 'Z'))
				{
					break;
				}
				if ((*(s + 3) == 0xA7) && (*(s + 4) == 0xBD) && (*(s + 5) == 0xA7))
				{
					break;
				}
				// 色指定でなければ通常の文字として扱う
				*p++ = *s++;
				*p++ = *s++;
				*p++ = *s++;
				continue;
			}
			*p++ = *s++;
			// UTF-8の2-3バイト目
			if (((*s >= 0x81) && (*s <= 0xBF)) && ((*(s+1) >= 0x81) && (*(s+1) <= 0xBF)))
			{
				*p++ = *s++;
				*p++ = *s++;
				continue;
			}
			// 不明な文字は1バイト文字として処理する
			continue;

		case 0xC0:
		case 0xC1:
		case 0xC3:
		case 0xC4:
		case 0xC5:
		case 0xC6:
		case 0xC7:
		case 0xC8:
		case 0xC9:
		case 0xCA:
		case 0xCB:
		case 0xCC:
		case 0xCD:
		case 0xCE:
		case 0xCF:
		case 0xD0:
		case 0xD1:
		case 0xD2:
		case 0xD3:
		case 0xD4:
		case 0xD5:
		case 0xD6:
		case 0xD7:
		case 0xD8:
		case 0xD9:
		case 0xDA:
		case 0xDB:
		case 0xDC:
		case 0xDD:
		case 0xDE:
		case 0xDF:
			*p++ = *s++;
			// UTF-8の2バイト文字 (§がある0xC2開始以外)
			if ((*s >= 0x80) && (*s <= 0xBF))
			{
				*p++ = *s++;
				continue;
			}
			// 不明な文字は1バイト文字として処理する
			continue;

		case 0xE0:
		case 0xE1:
		case 0xE2:
		case 0xE3:
		case 0xE4:
		case 0xE5:
		case 0xE6:
		case 0xE7:
		case 0xE8:
		case 0xE9:
		case 0xEA:
		case 0xEB:
		case 0xEC:
		case 0xED:
		case 0xEE:
			*p++ = *s++;
			// UTF-8の3バイト文字 (ｧがある0xEF開始以外)
			if (((*s >= 0x80) && (*s <= 0xBF)) && ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)))
			{
				*p++ = *s++;
				*p++ = *s++;
				continue;
			}
			// 不明な文字は1バイト文字として処理する
			continue;

		case 0xF0:
		case 0xF1:
		case 0xF2:
		case 0xF3:
		case 0xF4:
		case 0xF5:
		case 0xF6:
		case 0xF7:
			*p++ = *s++;
			// UTF-8の4バイト文字
			if (((*s >= 0x80) && (*s <= 0xBF)) && ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)) && ((*(s + 2) >= 0x80) && (*(s + 2) <= 0xBF)))
			{
			*p++ = *s++;
			*p++ = *s++;
			*p++ = *s++;
				continue;
			}
			// 不明な文字は1バイト文字として処理する
			continue;

		case 0x01:
		case 0x02:
		case 0x03:
		case 0x04:
		case 0x05:
		case 0x06:
		case 0x07:
		case 0x08:
		case 0x09:
		case 0x0A:
		case 0x0B:
		case 0x0C:
		case 0x0D:
		case 0x0E:
		case 0x0F:
		case 0x10:
		case 0x11:
		case 0x12:
		case 0x13:
		case 0x14:
		case 0x15:
		case 0x16:
		case 0x17:
		case 0x18:
		case 0x19:
		case 0x1A:
		case 0x1B:
		case 0x1C:
		case 0x1D:
		case 0x1E:
		case 0x1F:
			// 制御文字があれば直前で区切る
			break;

		case 0x20:
			// スペースがあれば直前で区切る
			break;

		default:
			*p++ = *s++;
			continue;
		}
		break;
	}

#if _INMM_LOG_OUTPUT
#ifndef _INMM_PERF_LOG
	*p = '\0';
	fprintf(fpLog, "CalcColorWordWrap: %s (%d) %s\n", (char *)lpBuffer, p - lpBuffer, (char *)lpString);
#endif
#endif

	return p - lpBuffer;
}

//
// 連英数字の分離禁則を計算する
//
// パラメータ
//   lpBuffer		文字列処理バッファ
//   lpString		対象文字列
// 戻り値
//   処理したバイト数
//
int CalcAlphanumericWordWrap(LPBYTE lpBuffer, LPCBYTE lpString)
{
	LPBYTE p = lpBuffer;
	LPCBYTE s = lpString;

	while (*s != '\0')
	{
		// 単語内の最大表示可能文字数を超える場合その場で区切る(保険)
		if (p - lpBuffer > nMaxWordChars)
		{
			break;
		}
		switch (*s)
		{
		case '0':
		case '1':
		case '2':
		case '3':
		case '4':
		case '5':
		case '6':
		case '7':
		case '8':
		case '9':
		case ',':
		case '.':
		case 'A':
		case 'B':
		case 'C':
		case 'D':
		case 'E':
		case 'F':
		case 'G':
		case 'H':
		case 'I':
		case 'J':
		case 'K':
		case 'L':
		case 'M':
		case 'N':
		case 'O':
		case 'P':
		case 'Q':
		case 'R':
		case 'S':
		case 'T':
		case 'U':
		case 'V':
		case 'W':
		case 'X':
		case 'Y':
		case 'Z':
		case 'a':
		case 'b':
		case 'c':
		case 'd':
		case 'e':
		case 'f':
		case 'g':
		case 'h':
		case 'i':
		case 'j':
		case 'k':
		case 'l':
		case 'm':
		case 'n':
		case 'o':
		case 'p':
		case 'q':
		case 'r':
		case 's':
		case 't':
		case 'u':
		case 'v':
		case 'w':
		case 'x':
		case 'y':
		case 'z':
		case 0x27: // #APOSTROPHE
			// 数字または数字を構成する記号ならば続けて処理する
			*p++ = *s++;
			continue;

		case '%':
			if ((*(s + 1) < '0') || (*(s + 1) > '9') || (*(s + 2) < '0') || (*(s + 2) > '9') || (*(s + 3) < '0') || (*(s + 3) > '9'))
			{
				// 数字の後ろの%はまとめて処理する
				*p++ = *s;
			}
			break;

		case 0x5C: // エスケープ記号 '\\'(shiftjis) or バックスラッシュ
			if (*(s + 1) == '%')
			{
				// 数字の後ろの%はまとめて処理する
				*p++ = *s++;
				*p++ = *s;
			}
			break;

		case 0xEF:
			switch (*(s + 1))
			{
			case 0xBD:
				switch (*(s + 2))
				{
				case 0x8C: // ，
				case 0x8E: // ．
					*p++ = *s++;
					*p++ = *s++;
					*p++ = *s++;
					continue;
				}
				break;

			case 0xBC:
				switch (*(s + 2))
				{
				case 0x85: // ％
					// 数字の後ろの%はまとめて処理する
					*p++ = *s++;
					*p++ = *s++;
					*p++ = *s;
					break;

				case 0x90: // ０
				case 0x91: // １
				case 0x92: // ２
				case 0x93: // ３
				case 0x94: // ４
				case 0x95: // ５
				case 0x96: // ６
				case 0x97: // ７
				case 0x98: // ８
				case 0x99: // ９
					// 2バイト数字
					*p++ = *s++;
					*p++ = *s++;
					*p++ = *s++;
					continue;
				}
				break;
			}
			break;
		}
		break;
	}

#if _INMM_LOG_OUTPUT
#ifndef _INMM_PERF_LOG
	*p = '\0';
	fprintf(fpLog, "CalcAlphanumericWordWrap: %s (%d) %s\n", (char *)lpBuffer, p - lpBuffer, (char *)lpString);
#endif
#endif

	return p - lpBuffer;
}

//
// 文字列長を取得する(最大指定あり)
//
// パラメータ
//   lpString		対象文字列
//   nMax			最大文字列長
// 戻り値
//   文字列長
//
int WINAPI strnlen0(LPCBYTE lpString, int nMax)
{
	if (lpString == NULL)
	{
		return 0;
	}

	LPCBYTE s = lpString;
	int len = 0;
	while (*s != '\0')
	{
		// 2バイトUtf-8の1バイト目
		if ((*s >= 0xC0) && (*s <= 0xDF))
		{
			// 2バイトUtf-8の2バイト目
			if ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF))
			{
				if (len + 1 >= nMax)
				{
					break;
				}
				s += 2;
				len += 2;
				continue;
			}
		}

		// 3バイトUtf-8の1バイト目
		if ((*s >= 0xE0) && (*s <= 0xEF))
		{
			// 3バイトUtf-8の2-3バイト目
			if (((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)) && ((*(s + 2) >= 0x80) && (*(s + 2) <= 0xBF)))
			{
				if (len + 1 >= nMax)
				{
					break;
				}
				s += 3;
				len += 3;
				continue;
			}
		}

		// 4バイトUtf-8の1バイト目
		if ((*s >= 0xF0) && (*s <= 0xF7))
		{
			// 4バイトUtf-8の2-4バイト目
			if (((*(s + 1) >= 0x80) && (*(s + 1) <= 0xBF)) && ((*(s + 2) >= 0x80) && (*(s + 2) <= 0xBF)) && ((*(s + 3) >= 0x80) && (*(s + 3) <= 0xBF)))
			{
				if (len + 1 >= nMax)
				{
					break;
				}
				s += 4;
				len += 4;
				continue;
			}
		}

		// それ以外の文字は1バイト文字として処理する
		if (len >= nMax)
		{
			break;
		}
		s++;
		len++;
	}

#if _INMM_LOG_OUTPUT
	fprintf(fpLog, "strnlen0: %s (%d) => %d\n", (char *)lpString, nMax, len);
#endif

	return len;
}

// winmm.dllのmciSendCommand APIへ転送
MCIERROR WINAPI _mciSendCommandA(MCIDEVICEID IDDevice, UINT uMsg, DWORD fdwCommand, DWORD dwParam)
{
	return lpfnMciSendCommandA(IDDevice, uMsg, fdwCommand, dwParam);
}

// winmm.dllのtimeBeginPeriod APIへ転送
MMRESULT WINAPI _timeBeginPeriod(UINT uPeriod)
{
	return lpfnTimeBeginPeriod(uPeriod);
}

// winmm.dllのtimeGetDevCaps APIへ転送
MMRESULT WINAPI _timeGetDevCaps(LPTIMECAPS ptc, UINT cbtc)
{
	return lpfnTimeGetDevCaps(ptc, cbtc);
}

// winmm.dllのtimeGetTime APIへ転送
DWORD WINAPI _timeGetTime(VOID)
{
	return lpfnTimeGetTime();
}

// winmm.dllのtimeKillEvent APIへ転送
MMRESULT WINAPI _timeKillEvent(UINT uTimerID)
{
	return lpfnTimeKillEvent(uTimerID);
}

// winmm.dllのtimeSetEvent APIへ転送
MMRESULT WINAPI _timeSetEvent(UINT uDelay, UINT uResolution, LPTIMECALLBACK lpTimeProc, DWORD dwUser, UINT fuEvent)
{
	return lpfnTimeSetEvent(uDelay, uResolution, lpTimeProc, dwUser, fuEvent);
}

//
// 初期化処理
//
void Init()
{
	// winmm.dllのAPI転送関数を設定する
	hWinMmDll = LoadLibrary("winmm.dll");
	lpfnMciSendCommandA = (LPFNMCISENDCOMMANDA)GetProcAddress(hWinMmDll, "mciSendCommandA");
	lpfnTimeBeginPeriod = (LPFNTIMEBEGINPERIOD)GetProcAddress(hWinMmDll, "timeBeginPeriod");
	lpfnTimeGetDevCaps = (LPFNTIMEGETDEVCAPS)GetProcAddress(hWinMmDll, "timeGetDevCaps");
	lpfnTimeGetTime = (LPFNTIMEGETTIME)GetProcAddress(hWinMmDll, "timeGetTime");
	lpfnTimeKillEvent = (LPFNTIMEKILLEVENT)GetProcAddress(hWinMmDll, "timeKillEvent");
	lpfnTimeSetEvent = (LPFNTIMESETEVENT)GetProcAddress(hWinMmDll, "timeSetEvent");

#if _INMM_LOG_OUTPUT
	fopen_s(&fpLog, "_inmm.log", "w");
#endif

	hDesktopDC = GetWindowDC(HWND_DESKTOP);

	SetFontDefault();
	LoadIniFile();
	InitFont();
	InitPalette();
	InitCoinImage();
}

//
// iniファイルから設定を読み込む
//
void LoadIniFile()
{
	// iniファイル名を設定する
	char szIniFileName[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, szIniFileName);
	lstrcat(szIniFileName, "\\_inmm.ini");

	// iniファイルを開く
	HANDLE hIniFile = CreateFile(szIniFileName, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hIniFile == NULL)
	{
		return;
	}

	// iniファイルのサイズを取得する
	DWORD dwIniFileSize = GetFileSize(hIniFile, NULL);
	if (dwIniFileSize == 0xFFFFFFFF)
	{
		return;
	}

	// iniファイル読み込み用のバッファを確保する
	char *lpszIniFileBuf = (char *)HeapAlloc(GetProcessHeap(), 0, dwIniFileSize + 1);
	if (lpszIniFileBuf == NULL)
	{
		return;
	}

	// iniファイルの内容を読み込む
	DWORD dwReadSize;
	if (ReadFile(hIniFile, lpszIniFileBuf, dwIniFileSize, &dwReadSize, NULL) == FALSE || dwReadSize != dwIniFileSize)
	{
		return;
	}
	lpszIniFileBuf[dwIniFileSize] = '\0';

	// iniファイルを閉じる
	CloseHandle(hIniFile);

	const int nLineBufSize = 512;
	char szLineBuf[nLineBufSize];
	const int nHexBufSize = 64;
	char szHexBuf[nHexBufSize];
	szHexBuf[0] = '0';
	szHexBuf[1] = 'x';
	int nMagicCode = 0;
	char *p = lpszIniFileBuf;

	while (p < lpszIniFileBuf + dwIniFileSize)
	{
		int len = ReadLine(p, szLineBuf, nLineBufSize);
		p += len;
		char *r = szLineBuf;
		// 空行/コメント行/セクション行は読み飛ばす
		if ((*r == '\0') || (*r == '#') || (*r == '['))
		{
			continue;
		}
		char *s = r;
		while ((*s != '=') && (*s != '\0'))
		{
			s++;
		}
		// 定義行の形式でない場合は無視する
		if (*s == '\0')
		{
			continue;
		}
		*s++ = '\0';
		// 色指定
		if ((*r == 'C') && (*(r + 1) == 'o') && (*(r + 2) == 'l') && (*(r + 3) == 'o') && (*(r + 4) == 'r') &&
			(*(r + 5) >= 'A') && (*(r + 5) <= 'Z') && (*s == '$'))
		{
			int nColorIndex = *(r + 5) - 'A';
			int c;
			lstrcpy(&szHexBuf[2], s + 1);
			StrToIntEx(szHexBuf, STIF_SUPPORT_HEX, &c);
			DWORD color = c & 0x00FFFFFF;
			colorTable[nColorIndex].r = (BYTE)(color & 0x000000FF);
			colorTable[nColorIndex].g = (BYTE)((color & 0x0000FF00) >> 8);
			colorTable[nColorIndex].b = (BYTE)((color & 0x00FF0000) >> 16);
		}
		else if (!lstrcmp(r, "MagicCode"))
		{
			nMagicCode = StrToInt(s);
			if ((nMagicCode < 0) || (nMagicCode >= nMaxFontTable))
			{
				nMagicCode = 0;
			}
		}
		else if (!lstrcmp(r, "Font"))
		{
			lstrcpy(fontTable[nMagicCode].szFaceName, s);
		}
		else if (!lstrcmp(r, "Height"))
		{
			fontTable[nMagicCode].nHeight = StrToInt(s);
		}
		else if (!lstrcmp(r, "Bold"))
		{
			fontTable[nMagicCode].nBold = StrToInt(s);
	}
		else if (!lstrcmp(r, "AdjustX"))
		{
			fontTable[nMagicCode].nAdjustX = StrToInt(s);
		}
		else if (!lstrcmp(r, "AdjustY"))
		{
			fontTable[nMagicCode].nAdjustY = StrToInt(s);
		}
		else if (!lstrcmp(r, "Coin"))
		{
			lstrcpy(szCoinFileName, s);
		}
		else if (!lstrcmp(r, "CoinAdjustX"))
		{
			nCoinAdjustX = StrToInt(s);
		}
		else if (!lstrcmp(r, "CoinAdjustY"))
		{
			nCoinAdjustY = StrToInt(s);
		}
		else if (!lstrcmp(r, "MaxLines"))
		{
			nMaxLines = StrToInt(s);
		}
		else if (!lstrcmp(r, "MaxWordChars"))
		{
			nMaxWordChars = StrToInt(s);
		}
	}
	
	// iniファイル読み込み用のバッファを解放する
	HeapFree(GetProcessHeap(), 0, lpszIniFileBuf);
}

//
// フォントの初期化
//
void InitFont()
{
	LOGFONT lf;
	char *p = (char *)&lf;
	for (int i = 0; i < sizeof(lf); i++)
	{
		*p = 0;
	}
	lf.lfWidth = 0;
	lf.lfEscapement = 0;
	lf.lfOrientation = 0;
	lf.lfItalic = FALSE;
	lf.lfUnderline = FALSE;
	lf.lfStrikeOut = FALSE;
	lf.lfCharSet = DEFAULT_CHARSET;
	lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
	lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
	lf.lfQuality = DEFAULT_QUALITY;
	lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;

	for (int i = 0; i < nMaxFontTable; i++)
	{
		if (fontTable[i].szFaceName[0] == '\0')
		{
			if (i == 0)
			{
				lstrcpy(fontTable[i].szFaceName, "MS UI Gothic");
			}
			else
			{
				continue;
			}
		}
		lf.lfHeight = (fontTable[i].nHeight != 0) ? fontTable[i].nHeight : 12;
		lf.lfWeight = (fontTable[i].nBold != 0) ? FW_BOLD : FW_NORMAL;
		lstrcpy(lf.lfFaceName, fontTable[i].szFaceName);
		fontTable[i].hFont = CreateFontIndirect(&lf);
	}
}

//
// フォントの初期値設定
//
void SetFontDefault()
{
	// arial10
	lstrcpy(fontTable[6].szFaceName, "MS UI Gothic");
	fontTable[6].nHeight = 7;
	fontTable[6].nAdjustY = 1;

	// arial8
	lstrcpy(fontTable[4].szFaceName, "MS UI Gothic");
	fontTable[4].nHeight = 5;
	fontTable[4].nAdjustY = 1;

	// eurostilecond10
	lstrcpy(fontTable[7].szFaceName, "MS UI Gothic");
	fontTable[7].nHeight = 7;
	fontTable[7].nAdjustY = 1;

	// eurostilecond12
	lstrcpy(fontTable[9].szFaceName, "MS UI Gothic");
	fontTable[9].nHeight = 9;
	fontTable[9].nAdjustY = 1;

	// eurostilecond14
	lstrcpy(fontTable[11].szFaceName, "MS UI Gothic");
	fontTable[11].nHeight = 12;

	// eurostilecond14bold
	lstrcpy(fontTable[10].szFaceName, "ＭＳ Ｐゴシック");
	fontTable[10].nHeight = 12;

	// eurostilecond18bold
	lstrcpy(fontTable[14].szFaceName, "MS UI Gothic");
	fontTable[14].nHeight = 14;
	fontTable[14].nBold = 1;
	fontTable[14].nAdjustY= 1;

	// eurostilecond20bold
	lstrcpy(fontTable[16].szFaceName, "ＭＳ Ｐゴシック");
	fontTable[16].nHeight = 12;
	fontTable[16].nAdjustY = 1;

	// eurostilecond24bold
	lstrcpy(fontTable[20].szFaceName, "ＭＳ Ｐゴシック");
	fontTable[20].nHeight = 12;
	fontTable[20].nAdjustY = 1;
}

//
// パレットの初期化
//
void InitPalette()
{
	LOGPALETTE *lpPalette = (LOGPALETTE *)HeapAlloc(GetProcessHeap(), 0, sizeof(LOGPALETTE) + sizeof(PALETTEENTRY) * (nMaxColorTable - 1));
	lpPalette->palVersion = 0x0300;
	lpPalette->palNumEntries = nMaxColorTable;
	for (int i = 0; i < nMaxColorTable; i++)
	{
		lpPalette->palPalEntry[i].peRed = colorTable[i].r;
		lpPalette->palPalEntry[i].peGreen = colorTable[i].g;
		lpPalette->palPalEntry[i].peBlue = colorTable[i].b;
		lpPalette->palPalEntry[i].peFlags = 0;
	}
	hPalette = CreatePalette(lpPalette);
	HeapFree(GetProcessHeap(), 0, lpPalette);
	SelectPalette(hDesktopDC, hPalette, FALSE);
	RealizePalette(hDesktopDC);
}

//
// コイン画像の初期化
//
void InitCoinImage()
{
	hBitmapCoin = (HBITMAP)LoadImage(NULL, szCoinFileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
	if (hBitmapCoin == NULL)
	{
		return;
	}
	hDCCoin = CreateCompatibleDC(hDesktopDC);
	hBitmapCoinPrev = (HBITMAP)SelectObject(hDCCoin, hBitmapCoin);
}

//
// 終了処理
//
void Terminate()
{
#if _INMM_LOG_OUTPUT
	fclose(fpLog);
#endif
	SelectObject(hDCCoin, hBitmapCoinPrev);
	DeleteObject(hBitmapCoin);
	DeleteDC(hDCCoin);
	DeleteObject((HGDIOBJ)hPalette);
	for (int i = 0; i < nMaxFontTable; i++)
	{
		if (fontTable[i].hFont != NULL)
		{
			DeleteObject((HGDIOBJ)fontTable[i].hFont);
		}
	}
	ReleaseDC(HWND_DESKTOP, hDesktopDC);
	FreeLibrary(hWinMmDll);
}

//
// 1行をバッファに読み込む
//
// パラメータ
//   p				文字列ポインタ
//   buf			文字列バッファ
//   len			文字列バッファの最大長
// 戻り値
//   1行のサイズ
//
int ReadLine(char *p, char *buf, int len)
{
	char *e = buf + len - 1;
	char *r = p;
	char *s = buf;
	while (s < e)
	{
		if (*r == '\0')
		{
			*s = '\0';
			return (r - p + 1);
		}
		if ((*r == '\r') || (*r == '\n'))
		{
			break;
		}
		*s++ = *r++;
	}
	*s = '\0';
	while (*r != '\0')
	{
		if ((*r != '\r') && (*r == '\n'))
		{
			break;
		}
		r++;
	}
	return (r - p + 1);
}
