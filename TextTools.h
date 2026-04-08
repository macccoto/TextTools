/*
 * TextTools.h
 * -----------
 * TextTools.dll の公開インターフェース定義
 *
 * C / C++ 両方から include できるようにしてある。
 * C++ からコンパイルする場合は extern "C" ブロックが有効になり、
 * 名前マングリング（C++ 独自の関数名装飾）が抑制される。
 */

#ifndef TEXTTOOLS_H
#define TEXTTOOLS_H

#include <wchar.h>  /* wchar_t の定義 */

/* ---- C++ から include された場合の名前マングリング抑制 ---- */
#ifdef __cplusplus
extern "C" {
#endif

/*
 * NormalizeTextW
 * --------------
 * Unicode（UTF-16）文字列を正規化する。
 *
 * 引数:
 *   src       入力文字列（NUL 終端 UTF-16）。NULL 不可。
 *   dest      出力バッファ（呼び出し元が確保）。NULL 不可。
 *   destChars バッファサイズ（wchar_t 単位、NUL 終端を含む）。
 *             例: 256 文字格納したい場合は 257 を渡す。
 *
 * 戻り値:
 *   > 0  : 書き込んだ文字数（NUL を除く）
 *   == 0 : 入力が空 / NULL / destChars <= 0
 *   バッファ不足の場合: 必要文字数（NUL 除く）を返す。
 *                       戻り値 >= destChars ならバッファ不足。
 *
 * 呼び出し側の判定パターン:
 *   ret = NormalizeTextW(src, dest, bufSize)
 *   if ret >= bufSize  → バッファ不足（ret + 1 サイズで再試行）
 *   if ret == 0        → 空文字列 or エラー
 *   if ret > 0 & ret < bufSize → 正常（dest[0..ret-1] が結果）
 */
int __stdcall NormalizeTextW(
    const wchar_t* src,
    wchar_t*       dest,
    int            destChars
);

#ifdef __cplusplus
}
#endif

#endif /* TEXTTOOLS_H */
