#################################################################
Excel + Pythonでデータ解析、集計が捗る！「Python in Excel」の紹介
#################################################################
.. raw:: html

   <a rel="license" href="http://creativecommons.org/licenses/by/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by/4.0/88x31.png" /></a><br /><small>This work is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by/4.0/">Creative Commons Attribution 4.0 International License</a>.</small>

はじめに
========

自己紹介
--------

* Ryuji Tsutsui @ryu22e
* `一般社団法人PyCon JP Association運営メンバー <https://www.pycon.jp/committee/members.html#ryuji-tsutsui>`_
* Python歴は13年くらい（主にDjango）
* Python Boot Camp、Shonan.py、GCPUG Shonanなどコミュニティ活動もしています
* 著書（共著）：『 `Python実践レシピ <https://gihyo.jp/book/2022/978-4-297-12576-9>`_ 』

今日話したいこと
----------------

* Excelの新機能「Python in Excel」とはどんな機能か
* どんな使い方をすると便利か
* ちょっとしたテクニックも紹介

どんな人に聞いてほしいか
------------------------

以下のいずれかに該当する人。

* Excelを使ってデータ解析、集計を行っている人
* Pythonを使ってデータ解析、集計を行っている人
* Excelの新機能に興味がある人

（※Pythonの経験がなくても大丈夫です）

このトークをやるモチベーション
------------------------------

* 私は、Python Boot CampというPythonチュートリアルイベントで講師をやっている
* Pythonに興味を持つ人を増やしたい
* Python in Excelを知ってもらうことで、Pythonの使い道をイメージできる人がいるかも？ という期待がある

トークの構成
------------

* Python in Excelの概要
* 使い方 & ちょっとしたテクニックを紹介

Python in Excelの概要
=====================

Python in Excelとは
-------------------

* ``=PY()`` というExcel関数を使って、ExcelのセルにPythonコードを埋め込める
* 「VBAをPythonで書けるようにする機能」ではない
* 2024年5月24日現在、Python in Excelはプレビュー段階の機能で、Windows版Excel（Excel for Windows）のみで利用可能

導入方法(1)
-----------

プレビュー段階の機能を利用するには、「Microsoft 365 Insider」のベータチャネルにサインアップする。

導入方法(2)
-----------

.. video:: _static/mp4/how-to-install-python-in-excel.mp4

どんな仕組みか
--------------

.. figure:: _static/img/python-in-excel-image.*
   :alt: PythonコードはMicrosoft Cloud上で実行される

   PythonコードはMicrosoft Cloud上で実行される

セキュリティについて(1)
-----------------------

他人が書いた不正なコードの実行を防ぐため、以下の制限がある。

* Excelの外にあるローカルリソースへのアクセス
* ネットワークアクセス
* 数式、グラフ、ピボットテーブル、マクロ、VBA コードなど、Excelブック内の他のプロパティへのアクセス

セキュリティについて(2)
-----------------------

Python実行環境のセキュリティアップデートはMicrosoft Cloudがやってくれる。

セキュリティについて(3)
-----------------------

Pythonコードが含まれているExcelファイルの入手元がインターネットまたは信頼されていないソースの場合、これを開くと保護ビューが有効になり、Pythonは実行されない。

Power Queryについて
-------------------

* Pythonからのネットワークアクセスができないので外部リソースの読み込みはできない
* その代わり、Power Queryを使って外部リソースのデータをセルに取り込んでからPythonで読み取ることはできる

使い方 & ちょっとしたテクニックを紹介
=====================================

以下についてデモをやります
--------------------------

* 簡単な計算
* 範囲選択のやり方
* 「出力形式」とは
* グラフの作成
* 「コアライブラリ」とは
* ちょっとしたテクニック

簡単な計算
----------

* ``=PY()`` というExcel関数を使って、セルにPythonコードを埋め込む
* セルの内容を読み取るには、 ``xl()`` 関数を使う

範囲選択のやり方
----------------

* ``xl("A1:A5")`` のようにセルの範囲を指定できる
* 範囲選択すると、PandasのDataFrameオブジェクトを取得できる

「出力形式」とは
----------------

=PY() Excel関数の出力形式には、以下の2種類がある。

Pythonオブジェクト（デフォルト）
    Pythonコードの実行結果をそのまま埋め込む出力形式。ひし形が2つ重なったアイコンが表示される。

Excelの値
    出力結果を人間に見せる際に使う出力形式。後述するグラフを作成する際にはこれを使う。

グラフの作成
------------

* データは「テーブル」にしておくと便利
* ``xl("テーブル名[#すべて]", headers=True)`` でテーブルのデータを取得
* ``sns.relplot(x="Month", y="Price", data=xl("テーブル1[#すべて]", headers=True), kind="line")`` でグラフを作成できる

「コアライブラリ」とは
----------------------

* Python in ExcelではAnacondaに同梱されているライブラリの一部が利用できる
* よく使うライブラリはimport文を書かずに使える
* これを「コアライブラリ」と呼ぶ

コアライブラリの一覧
--------------------

`Excel のオープンソース ライブラリと Python - Microsoft サポート <https://support.microsoft.com/ja-jp/office/excel-%E3%81%AE%E3%82%AA%E3%83%BC%E3%83%97%E3%83%B3%E3%82%BD%E3%83%BC%E3%82%B9-%E3%83%A9%E3%82%A4%E3%83%96%E3%83%A9%E3%83%AA%E3%81%A8-python-c817c897-41db-40a1-b9f3-d5ffe6d1bf3e>`_ を参照。

最後に
======

まとめ
------

ご清聴ありがとうございました
----------------------------

