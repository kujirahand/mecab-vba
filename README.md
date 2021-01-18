# mecab-vba

 - MeCab for Excel VBA (Windows)
 - 形態素解析器MeCabをVBAから使うライブラリ

## 利用準備

 - (1) [MeCab for Windows](https://taku910.github.io/mecab/#install-windows)をダウンロード
 - (2) MeCabをインストール。VBAなので、文字コードはShift_JISを推奨。Shift_JISでない場合、VBAで明示が必要。
 - (3) モジュールをインポート。Excel VBAに本アーカイブの mecab.basを追加。

## 利用例

```
Sub 形態素解析のテスト()

    ' MeCabインストール時の辞書文字コードを指定
    Call SetMeCabCharset("Shift_JIS")
    
    ' テスト用の文字列を指定
    TestStr = "探すのに時があり、諦めるのに時がある。"
    
    ' 文字列として結果を得る
    MsgBox MeCabExec(TestStr)
        
    ' シートに結果を入れる
    MeCabExecToSheet TestStr, Sheet1, 1

    ' MeCabItem配列に結果を得る
    Dim items() As MeCabItem
    items = MeCabExecToItems(TestStr)
    For i = 0 To UBound(items)
        Debug.Print items(i).表層形, items(i).ヨミ
    Next

End Sub
```


詳しくはsample.xlsmをご覧ください。


 