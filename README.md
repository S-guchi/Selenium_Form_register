# Selenium_Form_register
SeleniumVBAを使用。Excelの値を使用してForm入力

## 要件
ServiceNowっていうチケット管理システムへ色々な履歴を登録するツール  
流れとしては
1. ExcelにログインID、パスワード、登録事項を入力し実行ボタンをクリック
2. 随時登録処理を実行
3. 通信速度が遅いなどの理由で処理落ちしたら途中まで登録を終えているものに関しては登録済みとする。（再実行したときの２重登録を防ぐ）

## 補足
* この時は、Iframeっていう存在を知らず、要素が取れなくてかなりつまずいた    
* 拡張子vbとなっているが色をつけたかっただけ。  
* Chromeのバージョンが変わるとwebドライバーをダウンロードし直さなくてはいけないのでそこが辛い
