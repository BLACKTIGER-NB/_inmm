#include <stdio.h>
#include <Windows.h>
#include <WinGDI.h>
#include <ddraw.h>
#include <MMSystem.h>

#define _INMM_LOG_OUTPUT _DEBUG
//#define _INMM_PERF_LOG

static HDC hDesktopDC = NULL;
static HPALETTE hPalette = NULL;

// �t�H���g�e�[�u��
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

// �v���Z�b�g�J���[�e�[�u��
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
	{   0, 128,   0 }, // G
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

// �R�C���摜�\���p�̃������f�o�C�X�R���e�L�X�g�̃n���h��
static HDC hDCCoin = NULL;

// �R�C���摜�̃r�b�g�}�b�v�n���h��
static HBITMAP hBitmapCoin = NULL;

// �R�C���摜�ǂݍ��ݑO�̃r�b�g�}�b�v�n���h���ޔ��
static HBITMAP hBitmapCoinPrev = NULL;

// �R�C���摜�̃t�@�C����
static char szCoinFileName[MAX_PATH] = "Gfx\\Fonts\\Modern\\Additional\\coin.bmp";

// �R�C���摜�̕\���ʒu����
static int nCoinAdjustX = 0;
static int nCoinAdjustY = 0;

// �R�C���摜�̑傫��
static const int nCoinWidth = 12;
static const int nCoinHeight = 12;

// �ő�\���\�s��
static int nMaxLines = 84;

// �P����̍ő�\���\������
static int nMaxWordChars = 24;

// WINMM.DLL�̃��W���[���n���h��
static HMODULE hWinMmDll = NULL;

// �����񏈗��o�b�t�@
#define STRING_BUFFER_SIZE 1024

// ���O�o�͗p�t�@�C���|�C���^
#if _INMM_LOG_OUTPUT
static FILE *fpLog = NULL;
#endif

// ���J�֐�
int WINAPI GetTextWidth(LPCBYTE lpString, int nMagicCode, DWORD dwFlags);
void WINAPI TextOutDC0(int x, int y, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags);
void WINAPI TextOutDC1(LPRECT lpRect, int x, int y, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags);
void WINAPI TextOutDC2(LPRECT lpRect, int *px, int *py, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags);
int WINAPI CalcLineBreak(LPBYTE lpBuffer, LPCBYTE lpString);
int WINAPI strnlen0(LPCBYTE lpString, int nMax);

// �����֐�
static int GetTextWidthWord(LPCBYTE lpString, int nLen, int nMagicCode);
static void TextOutWord(LPCBYTE lpString, int nLen, LPPOINT lpPoint, LPRECT lpRect, HDC hDC, int nMagicCode, COLORREF color, DWORD dwFlags);
static void ShowCoinImage(LPPOINT lpPoint, HDC hDC);
static int CalcColorWordWrap(LPBYTE lpBuffer, LPCBYTE lpString);
static int CalcNumberWordWrap(LPBYTE lpBuffer, LPCBYTE lpString);
static void Init();
static void LoadIniFile();
static void InitFont();
static void InitPalette();
static void InitCoinImage();
static void Terminate();
static bool ReadLine(FILE *fp, char *buf, int len);

// winmm.dll��API�]���p
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
// DLL���C���֐�
//
// �p�����[�^
//   hInstDLL: DLL ���W���[���̃n���h��
//   fdwReason: �֐����Ăяo�����R
//   lpvReserved: �\��ς�
// �߂�l
//   ���������������TRUE��Ԃ�
//   
BOOL WINAPI DllMain(HINSTANCE hInstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	switch (fdwReason)
	{
	case DLL_PROCESS_ATTACH:
		Init();
		break;
	case DLL_THREAD_ATTACH:
		break;
	case DLL_THREAD_DETACH:
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
// �����񕝂��擾���� (�I���W�i���Ō݊�)
//
// �p�����[�^
//   lpString		�Ώە�����
//   nMagicCode		�t�H���g�w��p�}�W�b�N�R�[�h
//   dwFlags		�t���O(1�ő�����/2�ŉe�t��)
// �߂�l
//   ������
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

	while (*s != '\0')
	{
		switch (*s)
		{
		case 0x81:
			// ��
			if (*(s + 1) == 0x98)
			{
				// ���̌�ɉp�啶���������΃v���Z�b�g�̐F�w��
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
				// �����Ȃ�Ό��̐F�ɖ߂�
				if ((*(s + 2) == 0x81) && (*(s + 3) == 0x98))
				{
					if (s > p)
					{
						nWidth += GetTextWidthWord(p, s - p, nMagicCode);
					}
					s += 4;
					p = s;
					continue;
				}
				// ����3������0�`7�Ȃ�Β��l�̐F�w��
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
			// Shift_JIS��2�o�C�g��
			if (((*(s + 1) >= 0x40) && (*(s + 1) <= 0x7E)) || ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xFC)))
			{
				s += 2;
				continue;
			}
			// ����ȊO�Ȃ�ΒP�̂̕����Ƃ��Ĉ���
			s++;
			continue;

		case 0xA7: // �
			// ��̌�ɉp�啶���������΃v���Z�b�g�̐F�w��
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
			// ���Ȃ�Ό��̐F�ɖ߂�
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
			// ����3������0�`7�Ȃ�Β��l�̐F�w��
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
			// ����ȊO�Ȃ�ΒP�̂̕����Ƃ��Ĉ���
			s++;
			continue;

		case '%':
			// ����3������0�`7�Ȃ�Β��l�̐F�w��
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
			// ����ȊO�Ȃ�ΒP�̂̕����Ƃ��Ĉ���
			s++;
			continue;

		case '$':
			// �R�C���摜������Ε������Z����
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
			// �R�C���摜���Ȃ���ΒP�̂̕����Ƃ��Ĉ���
			s++;
			continue;

		case '\\': // �G�X�P�[�v�L��
			// �G�X�P�[�v�L���̌��%/�/$�������ΒP�̂̕����Ƃ��Ĉ���
			if (*(s + 1) == '%' || *(s + 1) == 0xA7 || *(s + 1) == '$')
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
			// �G�X�P�[�v�L���̌�Ɂ��������ΒP�̂̕����Ƃ��Ĉ���
			if ((*(s + 1) == 0x81) && (*(s + 2) == 0x98))
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
			// ����ȊO�Ȃ�ΒP�̂̕����Ƃ��Ĉ���
			s++;
			continue;

		case 0x0A:
		case 0x0D:
			// ���s�R�[�h
			if (s > p)
			{
				nWidth += GetTextWidthWord(p, s - p, nMagicCode);
			}
			s++;
			p = s;
			if (nWidth > nMaxWidth)
			{
				nMaxWidth = nWidth;
				nWidth = 0;
			}
			nLines++;
			// �ő�\���\�s���ɓ��B���Ă��Ȃ���Όv�Z���p������
			if (nLines < nMaxLines)
			{
				continue;
			}
			// �ő�\���\�s���ɓ��B������v�Z��ł��؂�
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
			// ���s�ȊO�̐���R�[�h�Ȃ�Γǂݔ�΂�
			if (s > p)
			{
				nWidth += GetTextWidthWord(p, s - p, nMagicCode);
			}
			s++;
			p = s;
			continue;

		case 0x82:
		case 0x83:
		case 0x84:
		case 0x85:
		case 0x86:
		case 0x87:
		case 0x88:
		case 0x89:
		case 0x8A:
		case 0x8B:
		case 0x8C:
		case 0x8D:
		case 0x8E:
		case 0x8F:
		case 0x90:
		case 0x91:
		case 0x92:
		case 0x93:
		case 0x94:
		case 0x95:
		case 0x96:
		case 0x97:
		case 0x98:
		case 0x99:
		case 0x9A:
		case 0x9B:
		case 0x9C:
		case 0x9D:
		case 0x9E:
		case 0x9F:
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
		case 0xEF:
		case 0xF0:
		case 0xF1:
		case 0xF2:
		case 0xF3:
		case 0xF4:
		case 0xF5:
		case 0xF6:
		case 0xF7:
		case 0xF8:
		case 0xF9:
		case 0xFA:
		case 0xFB:
		case 0xFC:
			// Shift_JIS��2�o�C�g��
			if (((*(s + 1) >= 0x40) && (*(s + 1) <= 0x7E)) || ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xFC)))
			{
				s += 2;
				continue;
			}
			// ����ȊO�Ȃ�ΒP�̂̕����Ƃ��Ĉ���
			s++;
			continue;

		default:
			// ����ȊO�Ȃ�ΒP�̂̕����Ƃ��Ĉ���
			s++;
			continue;
		}
		// �ő�\���\�s���ɓ��B�������Ɍv�Z��ł��؂邽�߂�break
		// ������ꍇ��switch���̒���continue�������ă��[�v�̐擪�֔�Ԃ���
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
	fprintf(fpLog, "GetTextWidth: %s (%d) %X => %d\n", lpString, nMagicCode, dwFlags, nMaxWidth);
#endif

	return nMaxWidth;
}

//
// �P��̕����񕝂��擾����
//
// �p�����[�^
//   lpString		�Ώە�����
//   nLen           ������(Byte�P��)
//   nMagicCode		�t�H���g�w��p�}�W�b�N�R�[�h
// �߂�l
//   ������
//
int GetTextWidthWord(LPCBYTE lpString, int nLen, int nMagicCode)
{
	SIZE size;
	GetTextExtentPoint32(hDesktopDC, (LPCSTR)lpString, nLen, &size);

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
// ��������o�͂��� (�I���W�i���Ō݊�)
//
// �p�����[�^
//   x				�o�͐�X���W
//   y				�o�͐�Y���W
//   lpString		�Ώە�����
//   lpDDS			DirectDraw�T�[�t�F�[�X�ւ̃|�C���^
//   nMagicCode		�t�H���g�w��p�}�W�b�N�R�[�h
//   dwColor		�F�w��(R-G-B��5bit-6bit-5bit��WORD�l��)
//   dwFlags		�t���O(1�ő�����/2�ŉe�t��)
//
void WINAPI TextOutDC0(int x, int y, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags)
{
	TextOutDC2(NULL, &x, &y, lpString, lpDDS, nMagicCode, dwColor, dwFlags);
}

//
// ��������o�͂��� (�N���b�s���O�Ή���)
//
// �p�����[�^
//   lpRect         �N���b�s���O��`(NULL�̏ꍇ�̓N���b�s���O�Ȃ�)
//   x				�o�͐�X���W
//   y				�o�͐�Y���W
//   lpString		�Ώە�����
//   lpDDS			DirectDraw�T�[�t�F�[�X�ւ̃|�C���^
//   nMagicCode		�t�H���g�w��p�}�W�b�N�R�[�h
//   dwColor		�F�w��(R-G-B��5bit-6bit-5bit��WORD�l��)
//   dwFlags		�t���O(1�ő�����/2�ŉe�t��)
//

void WINAPI TextOutDC1(LPRECT lpRect, int x, int y, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags)
{
	TextOutDC2(lpRect, &x, &y, lpString, lpDDS, nMagicCode, dwColor, dwFlags);
}

//
// ��������o�͂��� (�o�͈ʒu�X�V�Ή���)
//
// �p�����[�^
//   lpRect         �N���b�s���O��`(NULL�̏ꍇ�̓N���b�s���O�Ȃ�)
//   px				�o�͐�X���W�ւ̎Q��
//   py				�o�͐�Y���W�ւ̎Q��
//   lpString		�Ώە�����
//   lpDDS			DirectDraw�T�[�t�F�[�X�ւ̃|�C���^
//   nMagicCode		�t�H���g�w��p�}�W�b�N�R�[�h
//   dwColor		�F�w��(R-G-B��5bit-6bit-5bit��WORD�l��)
//   dwFlags		�t���O(1�ő�����/2�ŉe�t��)
//
void WINAPI TextOutDC2(LPRECT lpRect, int *px, int *py, LPCBYTE lpString, LPDIRECTDRAWSURFACE lpDDS, int nMagicCode, DWORD dwColor, DWORD dwFlags)
{
#if _INMM_LOG_OUTPUT
	fprintf(fpLog, "TextOutDC2 %s (%d,%d) %X %d %X %X\n", lpString, *px, *py, lpDDS, nMagicCode, dwColor, dwFlags);
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

	while (*s != '\0')
	{
		switch (*s)
		{
		case 0x81:
			// ��
			if (*(s + 1) == 0x98)
			{
				// ���̌�ɉp�啶���������΃v���Z�b�g�̐F�w��
				if ((*(s + 2) >= 'A') && (*(s + 2) <= 'Z'))
				{
					// �o�͂��镶���񂪂���Ε\������
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
				// �����Ȃ�Ό��̐F�ɖ߂�
				if ((*(s + 2) == 0x81) && (*(s + 3) == 0x98))
				{
					// �o�͂��镶���񂪂���Ε\������
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
				// ����3������0�`7�Ȃ�Β��l�̐F�w��
				if ((*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7') && (*(s + 4) >= '0') && (*(s + 4) <= '7'))
				{
					// �o�͂��镶���񂪂���Ε\������
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
			// Shift_JIS��2�o�C�g��
			if (((*(s + 1) >= 0x40) && (*(s + 1) <= 0x7E)) || ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xFC)))
			{
				s += 2;
				continue;
			}
			// ����ȊO�Ȃ�ΒP�̂̕����Ƃ��Ĉ���
			s++;
			continue;

		case 0xA7: // �
			// ��̌�ɉp�啶���������΃v���Z�b�g�̐F�w��
			if ((*(s + 1) >= 'A') && (*(s + 1) <= 'Z'))
			{
				// �o�͂��镶���񂪂���Ε\������
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
			// ���Ȃ�Ό��̐F�ɖ߂�
			if (*(s + 1) == 0xA7)
			{
				// �o�͂��镶���񂪂���Ε\������
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
			// ����3������0�`7�Ȃ�Β��l�̐F�w��
			if ((*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7'))
			{
				// �o�͂��镶���񂪂���Ε\������
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
			// ����ȊO�Ȃ�ΒP�̂̕����Ƃ��Ĉ���
			s++;
			continue;

		case '%':
			// ����3������0�`7�Ȃ�Β��l�̐F�w��
			if ((*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7') && (*(s + 3) >= '0') && (*(s + 3) <= '7'))
			{
				// �o�͂��镶���񂪂���Ε\������
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
			// ����ȊO�Ȃ�ΒP�̂̕����Ƃ��Ĉ���
			s++;
			continue;

		case '$':
			// �R�C���摜������Ε\������
			if (hBitmapCoin != NULL)
			{
				// �o�͂��镶���񂪂���Ε\������
				if (s > p)
				{
					TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
				}
				// �R�C���摜��\������
				ShowCoinImage(&pt, hDC);
				s++;
				p = s;
				continue;
			}
			// �R�C���摜���Ȃ���ΒP�̂̕����Ƃ��Ĉ���
			s++;
			continue;

		case '\\': // �G�X�P�[�v�L��
			// �G�X�P�[�v�L���̌��%/�/$�������ΒP�̂̕����Ƃ��Ĉ���
			if (*(s + 1) == '%' || *(s + 1) == 0xA7 || *(s + 1) == '$')
			{
				// �o�͂��镶���񂪂���Ε\������
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
			// �G�X�P�[�v�L���̌�Ɂ��������ΒP�̂̕����Ƃ��Ĉ���
			if ((*(s + 1) == 0x81) && (*(s + 2) == 0x98))
			{
				// �o�͂��镶���񂪂���Ε\������
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
			// ����ȊO�Ȃ�ΒP�̂̕����Ƃ��Ĉ���
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
			// �o�͂��镶���񂪂���Ε\������
			if (s > p)
			{
				TextOutWord(p, s - p, &pt, lpRect, hDC, nMagicCode, color, dwFlags);
			}
			// ����R�[�h��ǂݔ�΂�
			s++;
			p = s;
			continue;

		case 0x82:
		case 0x83:
		case 0x84:
		case 0x85:
		case 0x86:
		case 0x87:
		case 0x88:
		case 0x89:
		case 0x8A:
		case 0x8B:
		case 0x8C:
		case 0x8D:
		case 0x8E:
		case 0x8F:
		case 0x90:
		case 0x91:
		case 0x92:
		case 0x93:
		case 0x94:
		case 0x95:
		case 0x96:
		case 0x97:
		case 0x98:
		case 0x99:
		case 0x9A:
		case 0x9B:
		case 0x9C:
		case 0x9D:
		case 0x9E:
		case 0x9F:
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
		case 0xEF:
		case 0xF0:
		case 0xF1:
		case 0xF2:
		case 0xF3:
		case 0xF4:
		case 0xF5:
		case 0xF6:
		case 0xF7:
		case 0xF8:
		case 0xF9:
		case 0xFA:
		case 0xFB:
		case 0xFC:
			// Shift_JIS��2�o�C�g��
			if (((*(s + 1) >= 0x40) && (*(s + 1) <= 0x7E)) || ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xFC)))
			{
				s += 2;
				continue;
			}
			// ����ȊO�Ȃ�ΒP�̂̕����Ƃ��Ĉ���
			s++;
			continue;

		default:
			// ����ȊO�Ȃ�ΒP�̂̕����Ƃ��Ĉ���
			s++;
			continue;
		}
	}
	// �o�͂��镶���񂪂���Ε\������
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
// �P����o�͂���(�r���ŐF�ύX�͂Ȃ�)
//
// �p�����[�^
//   lpString		�Ώە�����
//   nLen           ������(Byte�P��)
//   lpPoint        �o�͐���W
//   lpRect         �N���b�s���O��`(NULL�̏ꍇ�̓N���b�s���O�Ȃ�)
//   hDC			�f�o�C�X�R���e�L�X�g�̃n���h��
//   color			�F�w��
//   dwFlags		�t���O(1�ő�����/2�ŉe�t��)
// 
void TextOutWord(LPCBYTE lpString, int nLen, LPPOINT lpPoint, LPRECT lpRect, HDC hDC, int nMagicCode, COLORREF color, DWORD dwFlags)
{
#if _INMM_LOG_OUTPUT
#ifndef _INMM_PERF_LOG
	fprintf(fpLog, "  TextOutWord %s %d (%d,%d) (%d,%d)-(%d,%d) %X %d %X %X\n", lpString, nLen, lpPoint->x, lpPoint->y,
		(lpRect != NULL) ? lpRect->left : 0, (lpRect != NULL) ? lpRect->top : 0, (lpRect != NULL) ? lpRect->right : 0, (lpRect != NULL) ? lpRect->bottom : 0,
		hDC, nMagicCode, (DWORD)color, dwFlags);
#endif
#endif

	int x = lpPoint->x + fontTable[nMagicCode].nAdjustX;
	int y = lpPoint->y + fontTable[nMagicCode].nAdjustY;

	// �e��t����ꍇ�͏o�͈ʒu�̉E���ɍ��ŏo�͂���
	bool bShadowed = ((dwFlags & 2) != 0);
	if (bShadowed)
	{
		SetTextColor(hDC, RGB(0, 0, 0));
		if (lpRect != NULL)
		{
			ExtTextOut(hDC, x + 1, y + 1, ETO_CLIPPED, lpRect, (LPCSTR)lpString, nLen, NULL);
		}
		else
		{
			TextOut(hDC, x + 1, y + 1, (LPCSTR)lpString, nLen);
		}
	}

	// �w��F�ŏo�͂���
	SetTextColor(hDC, color);
	if (lpRect != NULL)
	{
		ExtTextOut(hDC, x, y, ETO_CLIPPED, lpRect, (LPCSTR)lpString, nLen, NULL);
	}
	else
	{
		TextOut(hDC, x, y, (LPCSTR)lpString, nLen);
	}
	
	// �o�͈ʒu���X�V����
	SIZE size;
	GetTextExtentPoint32(hDC, (LPCSTR)lpString, nLen, &size);
	lpPoint->x += size.cx;
}

//
// �R�C���摜��\������
//
// �p�����[�^
//   lpPoint        �o�͐���W
//   hDC			�f�o�C�X�R���e�L�X�g�̃n���h��
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
// ���s�ʒu���v�Z���� (�֑������Ή���)
//
// �p�����[�^
//   lpBuffer		�����񏈗��o�b�t�@
//   lpString		�Ώە�����
// �߂�l
//   ���������o�C�g��
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
		// �s���֑�����/���䕶���̏���
		switch (*s)
		{
		case '(':
		case '<':
		case '[':
		case '`':
		case '{':
		case 0xA2: // �
			// �s���֑������Ȃ�Ύ��̕������ꏏ�ɏ�������
			*p++ = *s++;
			continue;

		case 0x81:
			*p++ = *s++;
			switch (*s)
			{
			case 0x4D: // �M
			case 0x65: // �e
			case 0x67: // �g
			case 0x69: // �i
			case 0x6B: // �k
			case 0x6D: // �m
			case 0x6F: // �o
			case 0x71: // �q
			case 0x73: // �s
			case 0x75: // �u
			case 0x77: // �w
			case 0x79: // �y
			case 0x83: // ��
			case 0x8F: // ��
			case 0x90: // ��
				// �s���֑������Ȃ�Ύ��̕������ꏏ�ɏ�������
				*p++ = *s++;
				continue;

			case 0x7B: // �{
			case 0x7C: // �|
			case 0x7D: // �}
				*p++ = *s++;
				// �A�����̕����֑�����
				len = CalcNumberWordWrap(p, s);
				p += len;
				s += len;
				break;

			case 0x98: // ��
				*p++ = *s++;
				// ���̌�ɉp�啶���������΃v���Z�b�g�̐F�w��
				if ((*s >= 'A') && (*s <= 'Z'))
				{
					// ��W�͔��F�Ȃ̂ŕ����֑��ΏۊO
					if (*s == 'W')
					{
						*p++ = *s++;
						// �F�w��̎��̕������ꏏ�ɏ�������
						continue;
					}
					// �F�w���̕����֑�����
					*p++ = *s++;
					len = CalcColorWordWrap(p, s);
					p += len;
					s += len;
					break;
				}
				// �����Ȃ�Ό��̐F�ɖ߂�
				if ((*s == 0x81) && (*(s + 1) == 0x98))
				{
					*p++ = *s++;
					*p++ = *s++;
					// �F�w��̎��̕������ꏏ�ɏ�������
					continue;
				}
				// ����3������0�`7�Ȃ�Β��l�̐F�w��
				if ((*s >= '0') && (*s <= '7') && (*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7'))
				{
					*p++ = *s++;
					*p++ = *s++;
					*p++ = *s++;
					// �F�w��̎��̕������ꏏ�ɏ�������
					continue;
				}
				break;

			default:
				// Shift_JIS��2�o�C�g��
				if (((*s >= 0x40) && (*s <= 0x7E)) || ((*s >= 0x80) && (*s <= 0xFC)))
				{
					*p++ = *s++;
					break;
				}
				// �s���ȕ�����1�o�C�g�����Ƃ��ď�������
				break;
			}
			break;

		case 0x82:
			*p++ = *s++;
			// 2�o�C�g����
			if ((*s >= 0x16) && (*s <= 0x25))
			{
				*p++ = *s++;
				// �A�����̕����֑�����
				len = CalcNumberWordWrap(p, s);
				p += len;
				s += len;
				break;
			}
			// Shift_JIS��2�o�C�g��
			if (((*s >= 0x40) && (*s <= 0x7E)) || ((*s >= 0x80) && (*s <= 0xFC)))
			{
				*p++ = *s++;
				break;
			}
			// �s���ȕ�����1�o�C�g�����Ƃ��ď�������
			break;

		case 0x87:
			*p++ = *s++;
			if (*s == 0x80) // ��
			{
				// �s���֑������Ȃ�Ύ��̕������ꏏ�ɏ�������
				*p++ = *s++;
				continue;
			}
			// Shift_JIS��2�o�C�g��
			if (((*s >= 0x40) && (*s <= 0x7E)) || ((*s >= 0x80) && (*s <= 0xFC)))
			{
				*p++ = *s++;
				break;
			}
			// �s���ȕ�����1�o�C�g�����Ƃ��ď�������
			break;

		case 0xA7: // �
			*p++ = *s++;
			// ��̌�ɉp�啶���������΃v���Z�b�g�̐F�w��
			if ((*s >= 'A') && (*s <= 'Z'))
			{
				// �W�͔��F�Ȃ̂ŕ����֑��ΏۊO
				if (*s == 'W')
				{
					*p++ = *s++;
					// �F�w��̎��̕������ꏏ�ɏ�������
					continue;
				}
				// �F�w���̕����֑�����
				*p++ = *s++;
				len = CalcColorWordWrap(p, s);
				p += len;
				s += len;
				break;
			}
			// ���Ȃ�Ό��̐F�ɖ߂�
			if (*s == 0xA7)
			{
				*p++ = *s++;
				// �F�w��̎��̕������ꏏ�ɏ�������
				continue;
			}
			// ����3������0�`7�Ȃ�Β��l�̐F�w��
			if ((*s >= '0') && (*s <= '7') && (*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7'))
			{
				*p++ = *s++;
				*p++ = *s++;
				*p++ = *s++;
				// �F�w��̎��̕������ꏏ�ɏ�������
				continue;
			}
			break;

		case '%':
			*p++ = *s++;
			// ����3������0�`7�Ȃ�Β��l�̐F�w��
			if ((*s >= '0') && (*s <= '7') && (*(s + 1) >= '0') && (*(s + 1) <= '7') && (*(s + 2) >= '0') && (*(s + 2) <= '7'))
			{
				*p++ = *s++;
				*p++ = *s++;
				*p++ = *s++;
				// �F�w��̎��̕������ꏏ�ɏ�������
				continue;
			}
			break;

		case '\\': // �G�X�P�[�v�L��
			*p++ = *s++;
			switch (*s)
			{
			case '$':
				// �G�X�P�[�v�L���̌��$�������ΒP�̂̍s���֑������Ƃ��Ĉ���
				*p++ = *s++;
				continue;

			case '%':
			case 0xA7: // �
				// �G�X�P�[�v�L���̌��%/��������ΒP�̂̕����Ƃ��Ĉ���
				*p++ = *s++;
				break;

			case 0x81:
				if (*(s + 1) == 0x98)
				{
					// �G�X�P�[�v�L���̌�Ɂ��������ΒP�̂̕����Ƃ��Ĉ���
					*p++ = *s++;
					*p++ = *s++;
				}
				else
				{
					// �P�̂�\�͍s���֑������Ƃ��Ĉ���
					continue;
				}
				break;

			default:
				// �P�̂�\�͍s���֑������Ƃ��Ĉ���
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
			*p++ = *s++;
			// �A�����̕����֑�����
			len = CalcNumberWordWrap(p, s);
			p += len;
			s += len;
			break;

		case 0x0A:
		case 0x0D:
			// �֑������̎��̉��s�R�[�h�͏������Ȃ�
			// ��������Ȃ��ƐF�w��̒���̉��s����������Ȃ�
			break;

		case 0x83:
		case 0x84:
		case 0x85:
		case 0x86:
		case 0x88:
		case 0x89:
		case 0x8A:
		case 0x8B:
		case 0x8C:
		case 0x8D:
		case 0x8E:
		case 0x8F:
		case 0x90:
		case 0x91:
		case 0x92:
		case 0x93:
		case 0x94:
		case 0x95:
		case 0x96:
		case 0x97:
		case 0x98:
		case 0x99:
		case 0x9A:
		case 0x9B:
		case 0x9C:
		case 0x9D:
		case 0x9E:
		case 0x9F:
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
		case 0xEF:
		case 0xF0:
		case 0xF1:
		case 0xF2:
		case 0xF3:
		case 0xF4:
		case 0xF5:
		case 0xF6:
		case 0xF7:
		case 0xF8:
		case 0xF9:
		case 0xFA:
		case 0xFB:
		case 0xFC:
			*p++ = *s++;
			// Shift_JIS��2�o�C�g��
			if (((*s >= 0x40) && (*s <= 0x7E)) || ((*s >= 0x80) && (*s <= 0xFC)))
			{
				*p++ = *s++;
				break;
			}
			// �s���ȕ�����1�o�C�g�����Ƃ��ď�������
			break;

		default:
			*p++ = *s++;
			break;
		}

		// �s���֑�����/�����֑������̏���
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
		case 0xA1: // �
		case 0xA3: // �
		case 0xA4: // �
		case 0xA5: // �
		case 0xA8: // �
		case 0xA9: // �
		case 0xAA: // �
		case 0xAB: // �
		case 0xAC: // �
		case 0xAD: // �
		case 0xAE: // �
		case 0xAF: // �
		case 0xB0: // �
		case 0xDE: // �
		case 0xDF: // �
			// �����s���֑������Ȃ�Έꏏ�ɏ�������
			continue;

		case '/':
			// ���������֑������Ȃ�΂��̎����܂߂Ĉꏏ�ɏ�������
			*p++ = *s++;
			continue;

		case '\\':
			// �����s���֑������Ȃ�Έꏏ�ɏ�������
			if (*(s + 1) == 0xA7) // �
			{
				continue;
			}
			break;

		case 0x81:
			switch (*(s + 1))
			{
			case 0x41: // �A
			case 0x42: // �B
			case 0x43: // �C
			case 0x44: // �D
			case 0x45: // �E
			case 0x46: // �F
			case 0x47: // �G
			case 0x48: // �H
			case 0x49: // �I
			case 0x4A: // �J
			case 0x4B: // �K
			case 0x4C: // �L
			case 0x52: // �R
			case 0x53: // �S
			case 0x54: // �T
			case 0x55: // �U
			case 0x56: // �V
			case 0x58: // �X
			case 0x5B: // �[
			case 0x5D: // �]
			case 0x60: // �`
			case 0x66: // �f
			case 0x68: // �h
			case 0x6A: // �j
			case 0x6C: // �l
			case 0x6E: // �n
			case 0x70: // �p 
			case 0x72: // �r
			case 0x74: // �t
			case 0x76: // �v
			case 0x78: // �x
			case 0x7A: // �z
			case 0x81: // ��
			case 0x84: // ��
			case 0x8E: // ��
			case 0x91: // ��
			case 0x93: // ��
				// �����s���֑������Ȃ�Έꏏ�ɏ�������
				continue;

			case 0x5C: // �\
			case 0x5E: // �^
			case 0x63: // �c
			case 0x64: // �d
				// ���������֑������Ȃ�΂��̎����܂߂Ĉꏏ�ɏ�������
				*p++ = *s++;
				*p++ = *s++;
				continue;
			}
			break;

		case 0x82:
			switch (*(s + 1))
			{
			case 0x9F: // ��
			case 0xA1: // ��
			case 0xA3: // ��
			case 0xA5: // ��
			case 0xA7: // ��
			case 0xC1: // ��
			case 0xE1: // ��
			case 0xE3: // ��
			case 0xE5: // ��
			case 0xEC: // ��
				// �����s���֑������Ȃ�Έꏏ�ɏ�������
				continue;
			}
			break;

		case 0x83:
			switch (*(s + 1))
			{
			case 0x40: // �@
			case 0x42: // �B
			case 0x44: // �D
			case 0x46: // �F
			case 0x48: // �H
			case 0x62: // �b
			case 0x83: // ��
			case 0x85: // ��
			case 0x87: // ��
			case 0x8E: // ��
			case 0x95: // ��
			case 0x96: // ��
				// �����s���֑������Ȃ�Έꏏ�ɏ�������
				continue;
			}
			break;

		case 0x87:
			if (*(s + 1) == 0x81) // ��
			{
				// �����s���֑������Ȃ�Έꏏ�ɏ�������
				continue;
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
	fprintf(fpLog, "CalcLineBreak: %s (%d) %s\n", lpBuffer, p - lpBuffer, lpString);
#endif

	return p - lpBuffer;
}

//
// �F�w���̕����֑����v�Z����
//
// �p�����[�^
//   lpBuffer		�����񏈗��o�b�t�@
//   lpString		�Ώە�����
// �߂�l
//   ���������o�C�g��
//
int CalcColorWordWrap(LPBYTE lpBuffer, LPCBYTE lpString)
{
	LPBYTE p = lpBuffer;
	LPCBYTE s = lpString;

	while (*s != '\0')
	{
		// �P����̍ő�\���\�������𒴂���ꍇ���̏�ŋ�؂�(�ی�)
		if (p - lpBuffer > nMaxWordChars)
		{
			break;
		}
		switch (*s)
		{
		case 0xA7: // �
			// �F�w��̒��O�ŋ�؂�
			if ((*(s + 1) >= 'A') && (*(s + 1) <= 'Z'))
			{
				break;
			}
			if (*(s + 1) == 0xA7)
			{
				break;
			}
			// �F�w��łȂ���Βʏ�̕����Ƃ��Ĉ���
			*p++ = *s++;
			continue;

		case 0x81:
			// ��
			if (*(s + 1) == 0x98)
			{
				// �F�w��̒��O�ŋ�؂�
				if ((*(s + 2) >= 'A') && (*(s + 2) <= 'Z'))
				{
					break;
				}
				if ((*(s + 2) == 0x81) && (*(s + 3) == 0x98))
				{
					break;
				}
				// �F�w��łȂ���Βʏ�̕����Ƃ��Ĉ���
				*p++ = *s++;
				*p++ = *s++;
				continue;
			}
			*p++ = *s++;
			// SHIFT_JIS��2�o�C�g��
			if (((*s >= 0x40) && (*s <= 0x7E)) || ((*s >= 0x80) && (*s <= 0xFC)))
			{
				*p++ = *s++;
				continue;
			}
			// �s���ȕ�����1�o�C�g�����Ƃ��ď�������
			continue;

		case 0x82:
		case 0x83:
		case 0x84:
		case 0x85:
		case 0x86:
		case 0x87:
		case 0x88:
		case 0x89:
		case 0x8A:
		case 0x8B:
		case 0x8C:
		case 0x8D:
		case 0x8E:
		case 0x8F:
		case 0x90:
		case 0x91:
		case 0x92:
		case 0x93:
		case 0x94:
		case 0x95:
		case 0x96:
		case 0x97:
		case 0x98:
		case 0x99:
		case 0x9A:
		case 0x9B:
		case 0x9C:
		case 0x9D:
		case 0x9E:
		case 0x9F:
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
		case 0xEF:
		case 0xF0:
		case 0xF1:
		case 0xF2:
		case 0xF3:
		case 0xF4:
		case 0xF5:
		case 0xF6:
		case 0xF7:
		case 0xF8:
		case 0xF9:
		case 0xFA:
		case 0xFB:
		case 0xFC:
			*p++ = *s++;
			// SHIFT_JIS��2�o�C�g��
			if (((*s >= 0x40) && (*s <= 0x7E)) || ((*s >= 0x80) && (*s <= 0xFC)))
			{
				*p++ = *s++;
				continue;
			}
			// �s���ȕ�����1�o�C�g�����Ƃ��ď�������
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
			// ���䕶��������Β��O�ŋ�؂�
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
	fprintf(fpLog, "CalcColorWordWrap: %s (%d) %s\n", lpBuffer, p - lpBuffer, lpString);
#endif
#endif

	return p - lpBuffer;
}

//
// �A�����̕����֑����v�Z����
//
// �p�����[�^
//   lpBuffer		�����񏈗��o�b�t�@
//   lpString		�Ώە�����
// �߂�l
//   ���������o�C�g��
//
int CalcNumberWordWrap(LPBYTE lpBuffer, LPCBYTE lpString)
{
	LPBYTE p = lpBuffer;
	LPCBYTE s = lpString;

	while (*s != '\0')
	{
		// �P����̍ő�\���\�������𒴂���ꍇ���̏�ŋ�؂�(�ی�)
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
			// �����܂��͐������\������L���Ȃ�Α����ď�������
			*p++ = *s++;
			continue;

		case '%':
			if ((*(s + 1) < '0') || (*(s + 1) > '9') || (*(s + 2) < '0') || (*(s + 2) > '9') || (*(s + 3) < '0') || (*(s + 3) > '9'))
			{
				// �����̌���%�͂܂Ƃ߂ď�������
				*p++ = *s++;
			}
			break;

		case '\\':
			if (*(s + 1) == '%')
			{
				// �����̌���%�͂܂Ƃ߂ď�������
				*p++ = *s++;
				*p++ = *s++;
			}
			break;

		case 0x81:
			switch (*(s + 1))
			{
			case 0x43: // �C
			case 0x44: // �D
				*p++ = *s++;
				*p++ = *s++;
				continue;

			case 0x93: // ��
				// �����̌���%�͂܂Ƃ߂ď�������
				*p++ = *s++;
				*p++ = *s++;
				break;
			}
			break;

		case 0x82:
			// 2�o�C�g����
			if ((*(s + 1) >= 0x16) && (*(s + 1) <= 0x25))
			{
				*p++ = *s++;
				*p++ = *s++;
				continue;
			}
			break;
		}
		break;
	}

#if _INMM_LOG_OUTPUT
#ifndef _INMM_PERF_LOG
	*p = '\0';
	fprintf(fpLog, "CalcNumberWordWrap: %s (%d) %s\n", lpBuffer, p - lpBuffer, lpString);
#endif
#endif

	return p - lpBuffer;
}

//
// �����񒷂��擾����(�ő�w�肠��)
//
// �p�����[�^
//   lpString		�Ώە�����
//   nMax			�ő啶����
// �߂�l
//   ������
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
		// Shift_JIS��1�o�C�g��
		if (((*s >= 0x81) && (*s <= 0x9F)) || ((*s >= 0xE0) && (*s <= 0xFC)))
		{
			// Shift_JIS��2�o�C�g��
			if (((*(s + 1) >= 0x40) && (*(s + 1) <= 0x7E)) || ((*(s + 1) >= 0x80) && (*(s + 1) <= 0xFC)))
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
		// ����ȊO�̕�����1�o�C�g�����Ƃ��ď�������
		if (len >= nMax)
		{
			break;
		}
		s++;
		len++;
	}

#if _INMM_LOG_OUTPUT
	fprintf(fpLog, "strnlen0: %s (%d) => %d\n", lpString, nMax, len);
#endif

	return len;
}

// winmm.dll��mciSendCommand API�֓]��
MCIERROR WINAPI _mciSendCommandA(MCIDEVICEID IDDevice, UINT uMsg, DWORD fdwCommand, DWORD dwParam)
{
	return lpfnMciSendCommandA(IDDevice, uMsg, fdwCommand, dwParam);
}

// winmm.dll��timeBeginPeriod API�֓]��
MMRESULT WINAPI _timeBeginPeriod(UINT uPeriod)
{
	return lpfnTimeBeginPeriod(uPeriod);
}

// winmm.dll��timeGetDevCaps API�֓]��
MMRESULT WINAPI _timeGetDevCaps(LPTIMECAPS ptc, UINT cbtc)
{
	return lpfnTimeGetDevCaps(ptc, cbtc);
}

// winmm.dll��timeGetTime API�֓]��
DWORD WINAPI _timeGetTime(VOID)
{
	return lpfnTimeGetTime();
}

// winmm.dll��timeKillEvent API�֓]��
MMRESULT WINAPI _timeKillEvent(UINT uTimerID)
{
	return lpfnTimeKillEvent(uTimerID);
}

// winmm.dll��timeSetEvent API�֓]��
MMRESULT WINAPI _timeSetEvent(UINT uDelay, UINT uResolution, LPTIMECALLBACK lpTimeProc, DWORD dwUser, UINT fuEvent)
{
	return lpfnTimeSetEvent(uDelay, uResolution, lpTimeProc, dwUser, fuEvent);
}

//
// ����������
//
void Init()
{
	// winmm.dll��API�]���֐���ݒ肷��
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

	LoadIniFile();
	InitFont();
	InitPalette();
	InitCoinImage();
}

//
// ini�t�@�C������ݒ��ǂݍ���
//
void LoadIniFile()
{
	char szIniFileName[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, szIniFileName);
	lstrcat(szIniFileName, "\\_inmm.ini");

	FILE *fp;
	fopen_s(&fp, szIniFileName, "r");
	if (fp == NULL)
	{
		return;
	}

	const int nLineBufSize = 512;
	char szLineBuf[nLineBufSize];
	int nMagicCode = 0;

	while (!feof(fp))
	{
		if (!ReadLine(fp, szLineBuf, nLineBufSize))
		{
			break;
		}
		char *p = szLineBuf;
		// ��s/�R�����g�s/�Z�N�V�����s�͓ǂݔ�΂�
		if ((*p == '\0') || (*p == '#') || (*p == '['))
		{
			continue;
		}
		char *s = p;
		while ((*s != '=') && (*s != '\0'))
		{
			s++;
		}
		// ��`�s�̌`���łȂ��ꍇ�͖�������
		if (*s == '\0')
		{
			continue;
		}
		*s++ = '\0';
		// �F�w��
		if (!strncmp(p, "Color", 5) && (*(p + 5) >= 'A') && (*(p + 5) <= 'Z') && (*s == '$'))
		{
			int nColorIndex = *(p + 5) - 'A';
			DWORD color = strtoul(s + 1, NULL, 16);
			colorTable[nColorIndex].r = (BYTE)(color & 0x000000FF);
			colorTable[nColorIndex].g = (BYTE)((color & 0x0000FF00) >> 8);
			colorTable[nColorIndex].b = (BYTE)((color & 0x00FF0000) >> 16);
		}
		else if (!lstrcmp(p, "MagicCode"))
		{
			nMagicCode = strtol(s, NULL, 0);
			if ((nMagicCode < 0) || (nMagicCode >= nMaxFontTable))
			{
				nMagicCode = 0;
			}
		}
		else if (!lstrcmp(p, "Font"))
		{
			lstrcpy(fontTable[nMagicCode].szFaceName, s);
		}
		else if (!lstrcmp(p, "Height"))
		{
			fontTable[nMagicCode].nHeight = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "Bold"))
		{
			fontTable[nMagicCode].nBold = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "AdjustX"))
		{
			fontTable[nMagicCode].nAdjustX = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "AdjustY"))
		{
			fontTable[nMagicCode].nAdjustY = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "Coin"))
		{
			lstrcpy(szCoinFileName, s);
		}
		else if (!lstrcmp(p, "CoinAdjustX"))
		{
			nCoinAdjustX = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "CoinAdjustY"))
		{
			nCoinAdjustY = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "MaxLines"))
		{
			nMaxLines = strtol(s, NULL, 0);
		}
		else if (!lstrcmp(p, "MaxWordChars"))
		{
			nMaxWordChars = strtol(s, NULL, 0);
		}
	}
	fclose(fp);
}

//
// �t�H���g�̏�����
//
void InitFont()
{
	LOGFONT lf;
	memset(&lf, 0, sizeof(lf));
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
		lf.lfHeight = fontTable[i].nHeight;
		lf.lfWeight = (fontTable[i].nBold != 0) ? FW_BOLD : FW_NORMAL;
		lstrcpy(lf.lfFaceName, fontTable[i].szFaceName);
		fontTable[i].hFont = CreateFontIndirect(&lf);
	}
}

//
// �p���b�g�̏�����
//
void InitPalette()
{
	LOGPALETTE *lpPalette = (LOGPALETTE *)malloc(sizeof(LOGPALETTE) + sizeof(PALETTEENTRY) * (nMaxColorTable - 1));
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
	free(lpPalette);
	SelectPalette(hDesktopDC, hPalette, FALSE);
	RealizePalette(hDesktopDC);
}

//
// �R�C���摜�̏�����
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
// �I������
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
// �t�@�C������1�s��ǂݍ��݁A�����̉��s�������폜����
//
// �p�����[�^
//   fp				�t�@�C���|�C���^
//   buf			������o�b�t�@
//   len			������o�b�t�@�̍ő咷
//
bool ReadLine(FILE *fp, char *buf, int len)
{
	if (fgets(buf, len, fp) == NULL)
	{
		return false;
	}
	char *p = strchr(buf, '\n');
	if (*p != NULL)
	{
		*p = '\0';
	}
	return true;
}
