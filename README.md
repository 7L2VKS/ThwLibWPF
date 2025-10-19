## Turbo HAMLOG/Win Input and Modify Window Access Class
Turbo HAMLOG/Win is the most popular logging software for amateur radio in Japan.<br/>
<br/>

## Turbo HAMLOG/Win 入力・修正ウィンドウ アクセスクラス

### 概要：
Turbo HAMLOG/Win の入力・修正ウィンドウに対してデータの配置、取得を行うためのアクセスクラス、およびそのクラスを利用したデモプログラムです。

アクセスクラスは Thw.cs にまとめています。その他のファイルはすべてデモを動かすためのファイルです。

### 必要パッケージ：
このプログラムをビルドするためには、NuGet から以下のパッケージのインストールが必要です。
- Prism.Core<br/>
  WPF の MVVM パターンに乗せるために使用
- Microsoft.Xaml.Behaviours.Wpf<br/>
  GUI の一部に WinForms コントロールを使用するために使用

### 使い方：
```
Thw thw = Thw.GetInstance();               // インスタンスの作成
thw.TargetWindow = Thw.WindowType.Input;   // 修正ウィンドウに対する操作の場合は.Editを指定
IntPtr hWnd = thw.SomeMethod();
```

### 資料：
Th527api に含まれる以下のファイルの情報をもとに作成しています。<br/>
- HamlogMs.txt

### ライセンス License
このプロジェクトは MIT ライセンスのもとで公開されています。This project is released under the MIT license. 
