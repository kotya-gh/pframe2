# TestLoader

## 概要

TestLoaderはコマンドを指定の順序で実行し、実行結果、コマンドのリターンコードを取得します。実行結果はファイルに自動的に記録されます。

## インストール法、始め方

TestLoaderフォルダ配下の「list.json」を設定後、「main.ps1」を実行します。

## 実行結果について

TestLoader実行後、次のような内容を標準出力します。
```
次のコマンドを実行します。: echo 'a'
コマンド実行後、次のメッセージが出力されます。: a
このコマンドは次のリターンコードを返します。: 0
コマンド実行開始時刻: 03/10/2023 18:42:54
コマンド実行終了時刻: 03/10/2023 18:42:54
コマンド実行結果
a
コマンドのリターンコード: 0
```
コマンド実行結果はテキストファイル、およびJSONファイルに記録されます。ファイル名は次のフォーマットで出力されます。
- 0001_hostname_YYYYMMDDHHMMSS_true.txt
- 0001_hostname_YYYYMMDDHHMMSS_true.json
- 0002_hostname_YYYYMMDDHHMMSS_false.txt
- 0002_hostname_YYYYMMDDHHMMSS_false.json

先頭4桁は後述する「testNo」で指定の項番です。
hostnameは「hostname」設定の文字列です。
true/falseは「returnCode」、「returnMsg」と実行結果を比較し、一致すればtrue、不一致の場合はfalseを付与します。

## list.json の設定

|変数名|説明|
|:--|:--|
|EvidenceHomeDir|実行結果を格納するディレクトリをフルパスで指定します。空の場合はスクリプトルート配下のfilesに保存します。|
|TestConfigure|実行するコマンドを格納する配列です。|
|testId|実行結果を格納するサブディレクトリです。EvidenceHomeDir配下に作成されます。|
|testItems|ホストごとに、実行したいコマンドを格納する配列です。|
|testNo|項番です。実行結果のファイル名の先頭４桁（０詰め）に付与されます。|
|hostname|コマンドを実行するホスト名です。スクリプト実行ホストと本設定が異なる場合、コマンド実行がスキップされます。|
|testCommands|個々のコマンドを格納する配列です。|
|order|testCommands内でのコマンド実行順序を指定します。|
|command|実行したいコマンドを指定します。|
|returnCode|期待するリターンコードを指定します。期待値と一致するかチェックし、結果を返します。|
|returnMsg|期待する出力メッセージを指定します。期待値と一致するかチェックし、結果を返します。空の場合はチェックをスキップします。|

## ライセンス

MIT