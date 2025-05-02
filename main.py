import streamlit as st
import pandas as pd
import io

st.title("Excelファイル比較アプリ")

st.write("""
このアプリは、2つのExcelファイルを比較し、指定した列の値が一致する行を認識します。
一致した行には 'Yes'、一致しない行には 'No' という結果を、1つ目のファイルに追加して出力します。
""")

# ファイルアップロード
st.header("1. Excelファイルのアップロード")
file1 = st.file_uploader("1つ目のExcelファイルを選択してください", type=["xlsx"])
file2 = st.file_uploader("2つ目のExcelファイルを選択してください", type=["xlsx"])

df1 = None
df2 = None

if file1 is not None:
    try:
        df1 = pd.read_excel(file1)
        # カラム名をすべて文字列に変換して表示・比較に使用する ★修正点★
        df1.columns = df1.columns.astype(str)
        st.write("1つ目のファイルのカラム:")
        st.write(df1.columns.tolist()) # 文字列に変換されたカラム名を表示
    except Exception as e:
        st.error(f"1つ目のファイルの読み込み中にエラーが発生しました: {e}")

if file2 is not None:
    try:
        df2 = pd.read_excel(file2)
        # カラム名をすべて文字列に変換して表示・比較に使用する ★修正点★
        df2.columns = df2.columns.astype(str)
        st.write("2つ目のファイルのカラム:")
        st.write(df2.columns.tolist()) # 文字列に変換されたカラム名を表示
    except Exception as e:
        st.error(f"2つ目のファイルの読み込み中にエラーが発生しました: {e}")

# 比較対象列の指定
st.header("2. 比較対象列の指定")
st.write("比較に使用する列名をカンマ区切りで入力してください。（例: カラム名A,カラム名B）")
st.write("※アップロード後に表示されるカラム名をご確認ください。") # 入力時の注意喚起

if df1 is not None and df2 is not None:
    # 入力は引き続き文字列として受け付ける
    cols_to_compare_str1 = st.text_input("1つ目のファイルの比較対象列名 (カンマ区切り)", key="cols1")
    cols_to_compare_str2 = st.text_input("2つ目のファイルの比較対象列名 (カンマ区切り)", key="cols2")

    # 入力された文字列をカンマで分割し、リストにする
    cols_to_compare1 = [col.strip() for col in cols_to_compare_str1.split(',') if col.strip()]
    cols_to_compare2 = [col.strip() for col in cols_to_compare_str2.split(',') if col.strip()]

    # バリデーション
    valid_cols = True
    if not cols_to_compare1 or not cols_to_compare2:
        valid_cols = False
        # st.warning("比較対象列を指定してください。") # ボタン押下時にまとめてチェック

    if valid_cols and len(cols_to_compare1) != len(cols_to_compare2):
         st.warning("1つ目と2つ目のファイルで指定する比較対象列の数が一致していません。")
         valid_cols = False

    if valid_cols:
        # 指定された列が各DataFrameに存在するかチェック
        # df.columnsは文字列になっているため、文字列として比較
        for col in cols_to_compare1:
            if col not in df1.columns:
                st.error(f"1つ目のファイルに指定された列 '{col}' が見つかりません。入力された列名が正確か、大文字・小文字、スペース等含めご確認ください。")
                valid_cols = False
        for col in cols_to_compare2:
            if col not in df2.columns:
                 st.error(f"2つ目のファイルに指定された列 '{col}' が見つかりません。入力された列名が正確か、大文字・小文字、スペース等含めご確認ください。")
                 valid_cols = False

    # 比較実行ボタン
    st.header("3. 比較実行")
    if st.button("比較を開始"):
        if df1 is not None and df2 is not None and valid_cols:
            st.write("比較を実行中です...")

            # 比較ロジック
            # 指定された列（既に文字列カラム名になっている）の値のタプルを作成して比較
            df2_compare_set = set(df2[cols_to_compare2].itertuples(index=False, name=None))

            # 1つ目のファイルの各行が2つ目のファイルの比較セットに含まれるかチェック
            df1['Comparison Result'] = df1[cols_to_compare1].apply(
                lambda row: 'Yes' if tuple(row) in df2_compare_set else 'No',
                axis=1
            )

            st.success("比較が完了しました！")

            # 結果のダウンロード
            st.header("4. 結果のダウンロード")
            output = io.BytesIO()
            try:
                # 結果を含むDataFrame（カラム名は文字列）をExcelファイルとして保存
                df1.to_excel(output, index=False)
                st.download_button(
                    label="結果ファイルをダウンロード (comparison_result.xlsx)",
                    data=output.getvalue(),
                    file_name="comparison_result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"結果ファイルの作成中にエラーが発生しました: {e}")

        else:
            # ボタン押下時のバリデーションメッセージを強化
            if df1 is None or df2 is None:
                st.warning("比較を開始するには、両方のファイルをアップロードしてください。")
            elif not cols_to_compare1 or not cols_to_compare2:
                 st.warning("比較を開始するには、比較対象列を指定してください。")
            elif valid_cols and len(cols_to_compare1) != len(cols_to_compare2):
                 st.warning("比較を開始するには、1つ目と2つ目のファイルで指定する比較対象列の数を一致させてください。")
            elif not valid_cols:
                 st.warning("指定された比較対象列に誤りがあります。アップロード後のカラム名を確認してください。")
