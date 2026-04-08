/*
 * TextTools.c
 * -----------
 * Excel VBA から呼び出す文字列正規化 DLL（最小版・Unicode 対応）
 *
 * ビルド方法:
 *   32bit: cl /LD /W4 TextTools.c /link /DEF:TextTools.def /OUT:TextTools32.dll
 *   64bit: cl /LD /W4 TextTools.c /link /DEF:TextTools.def /OUT:TextTools64.dll
 *
 * 注意: Visual Studio Developer Command Prompt で実行すること
 */

#include "TextTools.h"
#include <wchar.h>   /* wcslen, wchar_t */

/* ------------------------------------------------------------------ */
/* 内部ヘルパー関数                                                    */
/* ------------------------------------------------------------------ */

/*
 * isTrimChar:
 *   先頭・末尾トリミング対象の文字かどうかを返す
 *   対象: 半角スペース ' '、タブ '\t'、CR '\r'、LF '\n'
 */
static int isTrimChar(wchar_t c)
{
    return (c == L' ' || c == L'\t' || c == L'\r' || c == L'\n');
}

/*
 * isDeleteControl:
 *   削除すべき制御文字かどうかを返す
 *   0x00〜0x1F のうち、CR / LF / TAB を除いたもの
 *   （CR / LF / TAB は別ルールで処理するため除外）
 */
static int isDeleteControl(wchar_t c)
{
    return (c >= 0x01 && c <= 0x1F && c != L'\r' && c != L'\n' && c != L'\t');
    /* 0x00（NUL）は wcslen で長さ計算済みなので通常到達しないが、
       念のため 0x01 スタートにして NUL を正規文字として通過させない。
       実際には NUL が埋め込まれた文字列は VBA から来ない想定。    */
}

/* ------------------------------------------------------------------ */
/* メイン処理: NormalizeTextW                                          */
/* ------------------------------------------------------------------ */

/*
 * NormalizeTextW
 * --------------
 * 入力文字列 src を正規化して dest に書き込む。
 *
 * 正規化ルール:
 *   1. 先頭・末尾の ' ' '\t' '\r' '\n' を除去
 *   2. CRLF / CR / LF → 半角スペース 1 個
 *   3. TAB → 半角スペース 1 個
 *   4. 連続スペース → スペース 1 個に圧縮
 *   5. 制御文字（0x01〜0x1F、CR/LF/TAB 除く）→ 削除
 *
 * 戻り値:
 *   成功       : 書き込んだ文字数（終端 NUL を除く）
 *   バッファ不足: 必要文字数（終端 NUL を除く）を返す。dest には何も書かない
 *   引数エラー : 0
 */
int __stdcall NormalizeTextW(
    const wchar_t* src,   /* [in]  正規化したい文字列（NUL 終端） */
    wchar_t*       dest,  /* [out] 結果を受け取るバッファ */
    int            destChars /* [in] dest のバッファサイズ（wchar_t 単位、NUL 含む） */
)
{
    int      srcLen;       /* src の文字数（NUL 除く） */
    int      start;        /* トリミング後の開始インデックス */
    int      end;          /* トリミング後の終了インデックス（この位置を含む） */
    int      i;            /* ループカウンタ */
    int      prevWasSpace; /* 直前に空白を出力したかフラグ */
    int      requiredLen;  /* 出力に必要な文字数（NUL 除く） */
    int      outIdx;       /* dest への書き込み位置 */
    wchar_t  c;            /* 現在処理中の文字 */

    /* ---- 引数検証 ---- */
    if (src == NULL || dest == NULL || destChars <= 0) {
        return 0;
    }

    srcLen = (int)wcslen(src);

    /* ---- ステップ 1: 先頭・末尾トリミング ---- */
    start = 0;
    end   = srcLen - 1;

    /* 先頭をスキャンしてトリム文字をスキップ */
    while (start <= end && isTrimChar(src[start])) {
        start++;
    }
    /* 末尾をスキャンしてトリム文字をスキップ */
    while (end >= start && isTrimChar(src[end])) {
        end--;
    }

    /* トリム後に何も残らない場合 → 空文字列を返す */
    if (start > end) {
        if (destChars >= 1) {
            dest[0] = L'\0';
        }
        return 0;
    }

    /* ----------------------------------------------------------------
     * パス 1: 出力に必要な文字数をカウントする
     *         （バッファが足りるか先に確認するため）
     * ---------------------------------------------------------------- */
    requiredLen  = 0;
    prevWasSpace = 0;
    i = start;

    while (i <= end) {
        c = src[i];

        /* CRLF を 1 単位として処理（先に判定すること） */
        if (c == L'\r' && (i + 1 <= end) && src[i + 1] == L'\n') {
            if (!prevWasSpace) {
                requiredLen++;
                prevWasSpace = 1;
            }
            i += 2;  /* CR と LF の 2 文字分進める */
            continue;
        }

        /* 単独 CR / LF → スペース */
        if (c == L'\r' || c == L'\n') {
            if (!prevWasSpace) {
                requiredLen++;
                prevWasSpace = 1;
            }
            i++;
            continue;
        }

        /* TAB → スペース */
        if (c == L'\t') {
            if (!prevWasSpace) {
                requiredLen++;
                prevWasSpace = 1;
            }
            i++;
            continue;
        }

        /* 半角スペース → 圧縮 */
        if (c == L' ') {
            if (!prevWasSpace) {
                requiredLen++;
                prevWasSpace = 1;
            }
            i++;
            continue;
        }

        /* 削除対象の制御文字 → カウントしない（空白フラグも変えない） */
        if (isDeleteControl(c)) {
            i++;
            continue;
        }

        /* 通常の文字 */
        requiredLen++;
        prevWasSpace = 0;
        i++;
    }

    /* ---- バッファサイズチェック ----
     * destChars には NUL 終端の 1 文字分も含む必要がある。
     * つまり destChars > requiredLen でなければ書き込めない。
     * 足りない場合は必要文字数だけ返して終了（dest には書かない）。
     */
    if (destChars <= requiredLen) {
        return requiredLen;  /* 呼び出し元に「このサイズ以上が必要」を伝える */
    }

    /* ----------------------------------------------------------------
     * パス 2: 実際に dest へ書き込む
     *         （ロジックはパス 1 と完全に同じ）
     * ---------------------------------------------------------------- */
    outIdx       = 0;
    prevWasSpace = 0;
    i = start;

    while (i <= end) {
        c = src[i];

        /* CRLF */
        if (c == L'\r' && (i + 1 <= end) && src[i + 1] == L'\n') {
            if (!prevWasSpace) {
                dest[outIdx++] = L' ';
                prevWasSpace   = 1;
            }
            i += 2;
            continue;
        }

        /* CR / LF */
        if (c == L'\r' || c == L'\n') {
            if (!prevWasSpace) {
                dest[outIdx++] = L' ';
                prevWasSpace   = 1;
            }
            i++;
            continue;
        }

        /* TAB */
        if (c == L'\t') {
            if (!prevWasSpace) {
                dest[outIdx++] = L' ';
                prevWasSpace   = 1;
            }
            i++;
            continue;
        }

        /* スペース */
        if (c == L' ') {
            if (!prevWasSpace) {
                dest[outIdx++] = L' ';
                prevWasSpace   = 1;
            }
            i++;
            continue;
        }

        /* 制御文字削除 */
        if (isDeleteControl(c)) {
            i++;
            continue;
        }

        /* 通常文字 */
        dest[outIdx++] = c;
        prevWasSpace   = 0;
        i++;
    }

    /* NUL 終端 */
    dest[outIdx] = L'\0';

    return outIdx;  /* 書き込んだ文字数（NUL 除く）*/
}
