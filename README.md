# QRCodeActiveXControl

https://raw.githubusercontent.com/HiraokaHyperTools/QRCodeActiveXControl/master/how2.png

## 0.0.3

- エラーを出力し易い体質を改善しました。具体的には、エラーをコントロールを外側へ送信しないようにしました。代わりに ErrorMessage と HasErrors と Picture を設定します。
- 連結出力に対応しました。水平または垂直方法に並べることができます。
- いろいろ既定値を改めました。
- できるだけ QR コードを生成できるようにしました。たとえば、QRVersion の指定は固定のバージョンではなく、最小バージョンの指定となるように改めました。
