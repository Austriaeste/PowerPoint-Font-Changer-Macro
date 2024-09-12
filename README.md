# PowerPoint Font Changer Macro

PowerPointでプレゼン資料作成時にフォントを"MS 明朝"で指定されたので、マクロで一括処理できるようにしました。

## 動作確認

Microsoft 365で実施済みです。

## マクロの実行手順

1. **開発タブを有効にする**:
    - PowerPointを開き、`ファイル` > `オプション` > `リボンのユーザー設定`を選択します。
    - `開発`にチェックを入れて、`OK`をクリックします。

2. **マクロを追加する**:
    - `開発`タブをクリックし、`マクロ` > `Visual Basic`を選択します。
    - `挿入` > `モジュール`を選択し、上記のマクロコードを貼り付けます。
    - `ファイル` > `保存`を選択し、ファイル形式を`PowerPoint マクロ有効プレゼンテーション (*.pptm)`に設定して保存します。

3. **トラストセンターの設定**:
    - `ファイル` > `オプション` > `トラストセンター` > `トラストセンターの設定`を選択します。
    - `マクロの設定`を選択し、`すべてのマクロを有効にする`にチェックを入れます（セキュリティリスクがあるため、使用後は元に戻すことをお勧めします）。
    - `信頼できる場所`を選択し、マクロを含むファイルを保存するフォルダを追加します。

4. **マクロの実行**:
    - `開発`タブをクリックし、`マクロ`を選択します。
    - `ChangeFontToMSMincho`を選択し、`実行`をクリックします。

この手順に従うことで、PowerPointのプレゼン資料内のすべてのテキストフレーム、表、グループ化されたシェイプのフォントを一括で"MS 明朝"に変更することができます。


## メリット

テキストフレームのフォントを変更する処理を関数でまとめましたので、他のフォントに変更する場合も保守が楽だと思っています。

## マクロコード

```vba
Option Explicit
Sub ChangeFontToMSMincho()
    ' アクティブなプレゼンテーション内のすべてのスライド、シェイプ、表、グループ化されたシェイプのフォントを
    ' 指定したフォントに変更するマクロ

    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim row As Integer
    Dim col As Integer
    Dim grpShp As Shape

    ' 変更したいフォント名
    Dim targetFont As String
    targetFont = "ＭＳ 明朝"

    On Error GoTo ErrHandler

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            ' 通常のテキストフレーム、図形内のテキストフレーム、表のセル内のテキスト
            If shp.HasTextFrame Then
                With shp.TextFrame.TextRange.Font
                    .NameFarEast = targetFont
                End With
            ElseIf shp.HasTable Then
                Set tbl = shp.Table
                For row = 1 To tbl.Rows.Count
                    For col = 1 To tbl.Columns.Count
                        With tbl.Cell(row, col).Shape.TextFrame.TextRange.Font
                            .NameFarEast = targetFont
                        End With
                    Next col
                Next row
            ElseIf shp.Type = msoGroup Then ' グループ化されたシェイプ
                For Each grpShp In shp.GroupItems
                    If grpShp.HasTextFrame Then
                        With grpShp.TextFrame.TextRange.Font
                            .NameFarEast = targetFont
                        End With
                    End If
                Next grpShp
            End If
        Next shp
    Next sld

Exit Sub

ErrHandler:
    MsgBox "フォントの変更中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー説明: " & Err.Description
End Sub
