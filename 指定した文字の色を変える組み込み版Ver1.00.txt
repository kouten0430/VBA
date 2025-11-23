                    '---ここから指定した文字の色を変える処理---

                    'myRange（Rangeオブジェクト）、V（指定文字）、255（文字色）を必要に応じて変更して下さい
                    '指定文字が複数ある場合はInStrでヒットする最初の文字列のみ色が変わります

                    myRange.Characters(InStr(myRange.Value, V), Len(V)).Font.Color = 255

                    '---ここまで---
