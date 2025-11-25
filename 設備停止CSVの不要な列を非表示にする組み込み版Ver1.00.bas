    '---ここから不要な列を非表示にする処理---

    Columns(Cells.Find("調整期間　開始日時", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("調整期間　終了日時", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("取扱区分", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("部門区分", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("集約箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("決定箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("調整期間　毎連区分", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("停止設備および線路名　停止設備１", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("停止設備および線路名　停止設備２", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("停止設備および線路名　停止設備３", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("停止設備および線路名　停止設備４", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("停止設備および線路名　停止設備５", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("停止設備および線路名　停止設備６", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("停止設備および線路名　停止設備７", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("停止設備および線路名　停止設備８", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("停止設備および線路名　停止設備９", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("停止設備および線路名　停止設備１０", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("停止／充電区分", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("決定時注釈", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("入力区分", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("設備管理箇所確認　確認事項", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("通信チェック　２ルート停止有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("通信回線運用箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("関係箇所送付先　送付先１", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("関係箇所送付先　送付先２", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("関係箇所送付先　送付先３", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("関係箇所送付先　送付先４", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("関係箇所送付先　送付先５", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("関係箇所送付先　送付先６", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("関係箇所送付先　送付先７", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("関係箇所送付先　送付先８", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("関係箇所送付先　送付先９", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("関係箇所送付先　送付先１０", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("関係箇所送付先　送付先１１", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("配電線ＣＢ引出し", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("配電線接地", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("配電線切替先　有無設定", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("配電線切替先　切替先設備", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("配電Ｔｒ切替先　有無設定", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("配電Ｔｒ切替先　切替先設備", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("要求時作業条件　荒天中止", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("要求時作業条件　発雷中止", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("要求時作業条件　雨天中止", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("要求時作業条件　小雨決行", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("要求時作業条件　いっ水実施", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("要求時作業条件　貯渇水利用", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("決定時作業条件　荒天中止", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("決定時作業条件　発雷中止", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("決定時作業条件　雨天中止", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("決定時作業条件　小雨決行", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("決定時作業条件　いっ水実施", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("決定時作業条件　貯渇水利用", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電連絡責任者　氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電連絡責任者　ＴＥＬ", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("緊急復旧　可能／不可能", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("緊急復旧　復旧設備１", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("緊急復旧　設備１復旧時間", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("緊急復旧　復旧設備２", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("緊急復旧　設備２復旧時間", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("緊急復旧　復旧設備３", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("緊急復旧　設備３復旧時間", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("緊急時連絡ルート", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　接地有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　接続電気所１", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　電気所１代替有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　接続電気所２", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　電気所２代替有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　接続電気所３", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　電気所３代替有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　接続電気所４", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　電気所４代替有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　接続電気所５", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　電気所５代替有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　接続電気所６", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　電気所６代替有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　接続電気所７", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　電気所７代替有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　接続電気所８", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電接地　電気所８代替有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("需要停止　停止有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("需要停止　開始月日時分", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("需要停止　終了月日時分", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("需要停止　箇所１", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("需要停止　箇所２", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("需要停止　箇所３", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("需要停止　箇所４", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("需要停止　箇所５", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("需要停止　箇所６", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("変更時負担金　前日中止金額", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("変更時負担金　当日中止金額", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("停止電力（溢水）", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("Ｒｙ整定票　Ｎｏ．１", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("Ｒｙ整定票　Ｎｏ．２", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("Ｒｙ整定票　Ｎｏ．３", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("Ｒｙ整定票　Ｎｏ．４", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("Ｒｙ整定票　Ｎｏ．５", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("機器番号発効　Ｎｏ．１", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("機器番号発効　Ｎｏ．２", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("機器番号発効　Ｎｏ．３", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("機器番号発効　Ｎｏ．４", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("試験対応", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("系統運用細則", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("系統切替設備", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("運用停止設備", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("システム運用設定票Ｎｏ．", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ１　承認箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ１　承認日付", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ１　承認者氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ２　承認フラグ", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ２　承認箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ２　承認日付", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ２　承認者氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ３　承認フラグ", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ３　承認箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ３　承認日付", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ３　承認者氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ４　承認フラグ", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ４　承認箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ４　承認日付", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ４　承認者氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ５　承認フラグ", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ５　承認箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ５　承認日付", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("承認データ５　承認者氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電情報の停止／変動中給向有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("給電情報の停止／変動総制向有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("添付資料有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("中給件名削除有無", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（年間：要求）　最終登録日", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（年間：要求）　登録箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（年間：要求）　登録者氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（年間：集約）　最終登録日", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（年間：集約）　登録箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（年間：集約）　登録者氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（年間：調整）　最終登録日", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（年間：調整）　登録箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（年間：調整）　登録者氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（月間：要求）　最終登録日", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（月間：要求）　登録箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（月間：要求）　登録者氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（月間：集約）　最終登録日", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（月間：集約）　登録箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（月間：集約）　登録者氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（月間：調整）　最終登録日", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（月間：調整）　登録箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（月間：調整）　登録者氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（作業票：要求）　最終登録日", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（作業票：要求）　登録箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（作業票：要求）　登録者氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（作業票：集約）　最終登録日", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（作業票：集約）　登録箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（作業票：集約）　登録者氏名", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（作業票：調整）　最終登録日", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（作業票：調整）　登録箇所", LookAt:=xlWhole).Column).Hidden = True
    Columns(Cells.Find("登録データ（作業票：調整）　登録者氏名", LookAt:=xlWhole).Column).Hidden = True