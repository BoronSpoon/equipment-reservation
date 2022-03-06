## 0. 目次

1. はじめに
2. ソフトウェアの仕様
3. プログラムの内容の説明
4. 実装で気を付けた点
5. 簡易版操作マニュアル
6. ソースコード

## 1. はじめに

### このプログラムが必要な理由

* 実験装置の使用時間節約
  * 実験装置は使用開始時の立ち上げ操作と終了時の立ち下げ操作に最大でそれぞれ1時間程度かかるため、同じ日に複数人で実験装置を共有することで、立ち上げや立ち下げ操作を分担することができる。
* 使用時間の重複の防止
  * 実験装置の中には真空引きや冷却を行うのに最大で一晩近くかかるものがあり、使用時間がほかの人と重複すると実験予定が一日ずれてしまう。したがって、そのような事態を避けることができる。

### 想定される使われ方

* 装置の使用状況（予定の形で保存）をスマートフォンやPCのGoogleカレンダーアプリで確認する。その上でアプリで空いている時間帯に予定を追加する形で装置の予約を行う。使用状況には名前が書いてあり、予定の交渉はSlackや対面で行う。

## 2. ソフトウェアの仕様

### 概要

* Google Apps Scriptというサービスを用いる。これはGoogleのサービスをスクリプトでつなげられるサービスで、今回はサービスとしてSheets（表計算）とCalendar（カレンダー）を用いる。
* GAS上で用いる言語はJavascriptである。シンタックスとしてはV8（ChromeとNode.jsで用いられているランタイム）に対応したものを用いることができる。
* Sheetsはデータベースとして動作し、ユーザーの情報と設定内容を保存している。CalendarはUIとして動作し、予約の追加、変更、確認を行う。
* カレンダーにはユーザー側で操作するカレンダー（予約追加および変更用）とサーバー側で操作するカレンダー（予約確認用）が存在する。前者のカレンダーが操作されたとき、GASが後者のカレンダーに変更を反映する。以降、前者を書き込み専用カレンダーと呼び、後者を読み取り専用カレンダーと呼ぶ。

### 書き込み専用カレンダー（予約追加および変更用、ユーザー側で操作する）

予定を作成することで、使用する装置、装置の状態、使用時間を指定する。

* ユーザー毎にカレンダーが用意されている。ユーザー毎に用意されている理由はユーザーが自身に支給されたカレンダーに入力する場合、ユーザー名の入力を省略することができるためである。
* カレンダー名は`Write [User Name]`とする。ユーザー名がTana.Y1の場合はカレンダー名は`Write Tana.Y1`である。ユーザー名の命名の規則に関しては後に述べる。
* 予定の表記は`[device(小文字)] [state]`とする。CVDという装置を冷却する場合は`cvd cool`とする。deviceとstateの種類に関しては後に述べる。

#### カレンダーの名前(summary)とDescription

1. カレンダー名 `Write [User Name]`
2. カレンダーのDescription

   ```
   装置を予約する   
   Reserve devices   
   Formatting: [Device] [State]   
   Devices: rie, nrie(new RIE), cvd, ncvd, pvd, fts   
   States: evac(evacuation), use(or no entry), cool(cooldown), o2(O2 ashing)   
   
   ```

### 読み取り専用カレンダー（予約確認用、サーバー側で操作する）

他のユーザーの装置の使用状況を予定としてGoogleカレンダー上に表示する。

#### カレンダーの種類

カレンダーの種類は以下の2通りがある。

1. 全装置のカレンダーか使用する装置だけのカレンダーか
   1. 全装置の使用状況を含めた一つのカレンダー
   2. 自身が使用する装置をのみを表示したカレンダーを作成

#### 二種類のカレンダーを用意した理由

1. 全装置のカレンダーを表示するほうが良いと考えられる場合
   1. 全装置を扱う責任者(指導教員等)が全装置の使用状況を同時に把握したい場合。
2. 使用する装置のみを表示するカレンダーを作成するほうが良いと考えられる場合
   1. ユーザーは全装置を用いているわけではないため、自身の用いている装置のみの使用状況を把握したい場合。

#### 予定の表記

* 使用状況を示す予定はタイトルが`[ユーザー名] [装置名] [状態]`となっている。例えば、ユーザー名Tana.Y1がCVDという装置を冷却している場合、タイトルは`Tana.Y1 CVD cool`となっている。
* ユーザー毎にカレンダーが用意されている。ユーザー毎に用意されている理由はユーザーが自身に支給されたカレンダーに入力する場合、ユーザー名の入力を省略することができるためである。
* カレンダー名は`Read [user.name]`とする。ユーザーがTana.Y1の場合はカレンダー名は`Read Tana.Y1`である。全員分のカレンダーは名前は任意に決定できるが、今回はALL USERの省略として`Read ALL.U1`と設定する。

#### カレンダーの名前(summary)とDescription

1. カレンダー名 `Read [User Name]`
2. カレンダーのDescription

   ```
   装置の予約状況   
   schedule for selected devices   
   
   ```

### ユーザー名の表記

* ユーザー名は重複がないように以下のような形式になっている。数字は、姓名に重複のない場合は1であり、重複があった場合は後に作ったユーザーの数字を一つずつ増加させる。
  * `[姓の最初の4字(最初の文字は大文字)].[名の最初の一文字(大文字)][1,2,3,...]`

### ソフトウェアで用いる省略名称

スマートフォンのカレンダーアプリで表示することのできる文字数は制限されているため、装置名、状態、ユーザー名を省略表記で表示する。なお、装置の状態名は基本的に入力の簡単化のため、小文字になっている。装置名、ユーザー名はユーザーが入力することは基本的にはないため、大文字を用いている。

* Devices (**You can change the name of devices in Sheets**)
* | Device | Description | | ---- | ---- | | cmp | Chemical mechanical planarization | | rie | Reactive ion etching | | cvd | Chemical vapor deposition | | pvd | Physical vapor deposition | | euv | Extreme ultraviolet lithography | | diff | Diffusion | | cleg | Cleaning | | dicg | Dicing | | Pack | Packaging |
* States of equiment
* | State | 説明 | | ---- | ---- | | evac | evacuation | | use/[NO ENTRY] | Main operation | | cool | cooling |
* Users (**You can add users in Sheets**)
* | User Name | 名前 | | ---- | ---- | | Tana.Y1 | Tanaka Yuusuke | | Tana.Y2 | Tanabe Yuta | | Tana.S1 | Tanahashi Shion | | Chen.F1 | Chen Fan | | Zhu.Y1 | Zhu Yu |

## 3. プログラムの内容の説明

### プログラムの全体動作の説明

1. ユーザーの追加（初期設定）
   1. 新しいユーザーを追加する。
      1. Sheets上にユーザーの名前を記入する。
   2. 使用する装置を設定する。
      1. Sheets上で使用する装置にチェックを入れる。
   3. ユーザーのカレンダーを用意する。
      1. `Read [User Name]`と`Write [User Name]`の二つのカレンダーを用意
2. ユーザーがカレンダー上で予約を追加
3. GASがカレンダーの予定の変化を検知
   1. Calendarの予定編集時のトリガーを用いる。
   2. 変化した内容をみて、その装置を使用するユーザーの読み取り専用カレンダーに予定を表示する。
      1. 予定が追加された場合は読み取り専用カレンダーに予定を追加する。
      2. 予定が変更された場合は変更前の予定は消え、変更後の予定が読み取り専用カレンダーに表示されるようにする。
      3. 予定が削除された場合は読み取り専用カレンダーから予定を削除する。

### スクリプトの動作説明

* Standalone scriptを用いている。トリガーはinstallable triggerを用い、Sheets上での変更を検知する。Container-bound scriptを用いなかったのは、APIを呼び出す権限がなくカレンダー名を変更できないため、今回の目的を達成できないためである。
* カレンダーは事前に作成する必要がある。スクリプト一つにつき設定できるトリガーは20個であるため、学生19人分の読み取り専用、書き込み専用カレンダーを作成し、書き込み専用カレンダーにトリガーを設定した。残りの1個のトリガーはsheetsに設定した。卒業、入学のたびに卒業生のカレンダーを新入生が再利用することを想定している。これによりカレンダーが不足することはないと思われる。
* 全員の予定を確認することのできるカレンダーはALL.U1というユーザー名で作成した。このユーザーで書き込み専用カレンダーに予定を作成しない限り`Read ALL.U1`で全員分の予定を確認することができる。
* このスクリプトは次に示すようないくつかの役割を担っている。

1. ユーザーの追加、変更
2. ユーザーの追加、変更によるカレンダー名の変更
3. ユーザーの追加、変更による予定内のユーザー名の変更
4. ユーザーによる予定の追加、変更、削除
5. ユーザーが読み取り専用カレンダーで閲覧する装置の変更

#### 1. ユーザーの追加、変更

* Sheets上に名前を追加、変更する。名前の表記はローマ字表記で行い名、姓の順で表記し間にスペースをはさむ。
* Sheets上での変更をトリガーが検知する。
* 名前を名と姓に分解して、`[姓の最初の四文字].[名の最初の文字]`をUser Name 1として求める。
* すべてのユーザーのUser Name 1を見て上の行から順に重複防止用の番号をつけていき、User Name 2とする。これを以降ユーザー名(User Name)として採用する。User Name 2の表記は`[姓の最初の四文字].[名の最初の文字][1,2,3,...]`のようになる。

#### 2. ユーザーの追加、変更によるカレンダー名の変更

* ユーザーの追加、変更が行われた後にカレンダー名の変更が行われる。
* Sheetsの行ごとにカレンダーが二つ用意されている。これは読み取り専用と書き込み専用カレンダーに対応している。それぞれのカレンダーの名前をそれぞれ`Read [User Name]`と`Write [User Name]`に変更する。

#### 3. ユーザーの追加、変更による予定内のユーザー名の変更

* ユーザーの追加、変更が行われた後にカレンダー名の変更が行われ、そのあとに予定内のユーザー名の変更が行われる。

#### 4. ユーザーによる予定の追加、変更、削除

* Sync tokenという判別子を用いることで、変更された予定だけに処理を行う。これにより時間の節約が可能となる。
* 追加、変更、削除の3つの機能を最も単純に実装する予定の追加方法として読み取り専用カレンダーをゲストとして予定に追加する方法を用いた。これにより、ゲストに追加された読み取り専用カレンダーから予定を見ることができる。
  * 予定追加の場合は読み取り専用のカレンダーを予定のゲストとして追加する形で予定を追加する。
  * 予定変更の場合は、読み取り専用のカレンダーはすでに予定のゲストとして追加されてているため、もう一度ゲストとして追加しても変化はない。また、変更は自動的に読み取り専用カレンダーに適用される。
  * 予定削除の場合は、読み取り専用のカレンダーがゲストとして参加している元の予定が削除されるとその予定は読み取り専用のカレンダーから削除される。

#### 5. ユーザーが読み取り専用カレンダーで閲覧する装置の変更

* 閲覧する装置を変更した場合は、過去30日間にさかのぼり、閲覧するとして設定した装置の予定を追加する。前に閲覧していたが、今回閲覧しないと設定した装置に関しては予定を削除する。予定の追加、削除に関しては前と同様にゲストとして読み取り専用カレンダーに参加させることで追加する。

## 4. 実装で気を付けた注意点

#### 説明が簡潔なコーディング時の注意点

1. ローカル変数の長さ
   1. for文のインデックスでは`i`という短い変数を用い、for文以外ではわかりやすさを重視して長い変数を用いた。なお、`row, column`をインデックスとして用いた方がわかりやすいところがあったため、そこではfor文内であっても`i`以外の変数を用いた。
2. if文の入れ子を深くしすぎないために、else-ifを可能な限り用いた。
3. for文のidiomを以下のように統一した。

   ```
   for (var i = 0; i < maxIndex; i++){
      const item = items[i];
   }
   
   ```
4. 当たり前のことを表現するコメントを書かないように注意した。コメントが複雑になりすぎる場合には元のコードを簡単化するように意識した。また、magicコメントは書かないように注意した。
5. コメントはすべて英語で表記した。

#### javascriptの記法

https://google.github.io/styleguide/jsguide.html (以下jsguideと呼ぶ)の内容にもとづき、スクリプト中での記法を統一した。この中でも特に重要な点を示す。

1. スタイル
   1. 文字コードはUTF-8を用いた。
   2. `if, else, for, do, while`文においてブレースを必ず用い、K&Rスタイルで記述した。以下にコード中の例を示す。

      ```
      for (var column = 8; column < lastColumn+1; column++) { 
         sheet.getRange(cell.getRow(), column).insertCheckboxes();  
      }
      
      ```
   3. インデントはスペース2つ分とした。
   4. ステートメントの後にはセミコロンをつけた。
   5. コメントはVisual Studio Code(以下ではvscodeと表記する)での書きやすさの観点から複数行にわたる場合でも`//`を用いる(この記法はjsguideで許されている)。vscodeでは複数列を同時に編集できるため、その複数列が同じ文字列である方が編集がしやすい。`/* * */`の記法では全ての列が同じ文字でないため、編集時の手間が増えてしまう。

      ```
      // [全ての列が同じ文字である例]
      //
      //
      //
      
      /* [全ての列が同じ文字でない例]
      *
      *
      */
      
      ```
2. ローカル変数
   1. ローカル変数は`const`で宣言した。ただ、for文のインデックスや値が変化する変数、`const`のスコープ外で参照したい変数は`var`で宣言した。
   2. ローカル変数は必要とされる場所に一番近いところで宣言し、最も早く初期化するように意識した。
3. Arrayリテラル
   1. `const a1 = new Array(x1, x2, x3);`表記は禁止されているため、許可されている`const a1 = [x1, x2, x3];`表記を用いた。
4. オブジェクトリテラル
   1. 以下のようなkeyのクオート有り、無し表記の混在は禁止されているため、keyはすべてクオート無しで統一した。

      ```
      {
         width: 42, // struct-style unquoted key
         'maxWidth': 43, // dict-style quoted key
      }
      
      ```
5. Stringリテラル
   1. Stringリテラルはダブルクオート(`"`)ではなくシングルクオート(`'`)で囲む形で統一した。
   2. 複数行にわたる場合は以下のコード中の例のように`+`で結合した。

      ```
      const description = '装置を予約する\n' +
      'Reserve devices\n' +
      'Formatting: [Device] [State]\n' +
      'Devices: rie, nrie(new RIE), cvd, ncvd(new CVD), pvd, fts\n' +
      'States: evac(evacuation), use(or no entry), cool(cooldown), o2(RIE O2 ashing)\n';  
      
      ```
6. 制御文
   1. jsguide中ではfor文の構文として`for...of`文を用いることが推奨されていたが、`for...of`文を使える場面と使えない場面が混在しており、コーディング中に混乱が生じたため、`for (var i = 0; i < maxIndex; i++)`の記法に統一した。
   2. 比較演算子には`===`,`!==`を用いた。なお、`null`との比較に限り`==`を用いた。
7. 変数の命名規則
   1. キャメルケースをもちいた。
      1. クラス、Enumはアッパーキャメルケース(`UpperCamelCase`)を用いた。
      2. メソッド、パラメータ、ローカル変数にはローワーキャメルケース(`lowerCamelCase`)を用いた。
   2. 省略名称を用いてコード長を短くすることよりも新しくコードを読む人が理解しやすい変数名を用いた。
   3. 省略名称に関しては`URL`のように一般的に知られているものは用いたが、`wgc`のようなわからないものは用いていない。また、`pc`のように複数の選択肢がある場合も用いていない。

#### デバッグ時

1. gitでバージョン管理を行った。また、GitHubにprivateリポジトリを作成しデータの破損による影響を少なくした。コード自体のバックアップは定期的に行った。
2. テストに関しては次節で述べる。
3. テストにおいて処理時間は一回当たり15秒以内であったため、パフォーマンスの改善は必要ないと判断した。

#### テスト

用いたテストケースを以下で説明する。

##### Google Calendar上

- 予定
   - [x] 作成
   - 変更
      - [x] 状態(use、vac等)
      - [x] 日時
   - [x] 削除
   - [x] 複製
- 全日予定(一日中にわたる予定)
   - [x] 作成
   - 変更
      - [x] 状態(use、vac等)
      - [x] 日時
   - [x] 削除
   - [x] 複製
- 定期的な予定
   - [x] 作成
   - 変更
      - [x] 状態(use、vac等)
      - [x] 日時
   - [x] 削除
   - [x] 既に存在する予定を定期的な予定に変更する。

##### Google Sheets上
- users sheet
   - [x] 使用する装置の変更(使用する装置のチェックマークを入れたり外したりする)
   - ユーザー名(`User Name`)の変更
      - [x] カレンダー名`Read [User Name]`、`Write [User Name]`が変更される事の確認
      - [x] イベントの名前にユーザー名(`User Name`)が入っているため、イベント名が変更される事の確認
   - [x] カレンダーのリンクをクリックしてリンクの正常性の確認
- properties sheet
   - [x] 装置名の変更
   - [x] 実験条件の変更
   - [x] sheetのリンクをクリックしてリンクの正常性の確認
- equipment sheets
   - [x] 実験条件変更が反映されるか確認
   - [x] startTime, name, stateを変更できるか確認
   - 新しい予定を作る
      - [x] オンライン状態で作る
      - [x] オフライン状態で作り、オンラインにする
- backup
   - [x] イベント
   - [ ] daily logging
- その他
   - Sheetsをシェア
      - [x] Sheetsの所有者以外のユーザーによる操作でのtriggerの発火の確認。
      - [x] Sheetsの所有者以外のユーザーによる操作可能範囲の設定の確認

#### installable triggerのquotaの確認

Google Apps Scriptには以下のquotaが設定されている。19人のユーザーがいたとしても十分な実行回数が確保できると考える。

1. カレンダーの予約の作成：一日当たり5,000回実行可能
2. triggerの実行時間：一日当たり90分間実行可能
   1. 1回の実行時間を35秒とすると一日当たり150回実行可能

##### 実行時間

1. 予定作成 (max 35s @ 17 users) 2. 2人:9\~10s, 13人:12\~25s -> 18人14\~32s
2. 閲覧する装置の変更
   1. 17ユーザーで3\~5分ほど

#### Google Sheetsのサイズ

Workbook内のすべてのSheetの合計Cell数は最大で10000000=1e7

* variables
  * userCount = 17 users (18 when including the all event user)
  * equipmentCount = 50 equipments
  * experimentConditionCount = 20 experiment conditions for a single equipment
  * experimentConditionRows = 5000 rows in experiment condition
  * finalLoggingRows = 1000000 rows in final logging

##### 1. configSpreadsheet

* users sheet
  * (equipmentCount+9) columns \* userCount+2 rows = 5100 cells

|||||||||||
|-|-|-|-|-|-|-|-|-|-|-|
|Full Name|Last Name|First Name|User Name 1|User Name 2|Read Cal Id|Write Cal Id|Read Cal URL|Write Cal URL|**equipmentCount cols**|
|**userCount+1 rows**||||||||||

* properties sheet
  * (experimentConditionCount+3) columns \* (equipmentCount+1) rows = 1173 cells

|||||
|-|-|-|-|
|equipmentName|sheetName|sheetUrl|**experimentConditionCount cols**|
|**equipmentCount rows**||||

* 別々のworkbookにすると50個のWorkbookにTriggerを設定する必要があり、20 Trigger/User/Appの制限を超える
  * そのため、同じworkbook中の別のsheetにした。
* equipment sheet
  * (12+experimentConditionCount) columns \* (experimentConditionRows+1) rows \* 50 = 8.0e6 < 1e7 cells
* allEquipment sheet
  * 5 columns\* (1*+equipmentCount\*experimentConditionRows) rows = 1.25e6

||||||||||
|-|-|-|-|-|-|-|-|-|-|-|
|start Time|end Time|name|equipment|status|description|is AllDay Event|is Recurring Event|action|executionTime|id|eventExists|
|**experimentConditionRows rows**|||||||||

||||||
|-|-|-|-|-|
|startTime|id|state|{sheetName}!R{row}C12|eventExistsInRow|

##### 3. loggingSpreadsheet

* finalLog sheet
  * 8 columns \* finalLoggingRows rows = 8e6 < 1e7 cells

|||||||||
|-|-|-|-|-|-|-|-|-|-|
|startTime|endTime|name|equipment|status|description|isAllDayEvent|isRecurringEvent|
|**finalLoggingRows rows**||||||||

## 5. 簡易版操作マニュアル

### ユーザーの追加(初期設定)

1. 新しいユーザーを追加する。
   1. Sheets上にユーザーの名前を記入する(Sheetsの画像の➀に対応)。
2. 使用する装置を設定する。
   1. Sheets上で使用する装置にチェックを入れる(Sheetsの画像の➂に対応)。
3. ユーザーのカレンダーを用意する。
   1. `Read`(読み取り専用)と`Write`(書き込み専用)の二つのカレンダーのリンクが用意されている(Sheetsの画像の➁に対応)。
   2. マウスカーソルをリンク上に移動してカレンダー上に追加する。 ![](%22./pics/sheets.PNG%22)

### 予定の追加

1. Calendar上で予定を追加する。
   1. 予定のタイトルには`[装置名] [状態]`を記入する(Calendarの画像の➀に対応)。装置名、状態は前節の内容を参考にする。
   2. 追加するカレンダーは`Write [User Name]`を選択する(Calendarの画像の➁に対応)。 ![](%22./pics/calendar.PNG%22)

<script type="text/javascript" src="http://cdn.mathjax.org/mathjax/latest/MathJax.js?config=TeX-AMS-MML_HTMLorMML"></script> <script type="text/x-mathjax-config"> MathJax.Hub.Config({ tex2jax: {inlineMath: [['$', '$']]}, messageStyle: "none" }); </script>
