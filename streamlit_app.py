def convert_df_to_excel(files_data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Data")
        writer.sheets["Data"] = worksheet

        # フォント設定用のフォーマットを作成
        cell_format = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12})
        border_format = workbook.add_format({'font_name': 'Times New Roman', 'font_size': 12, 'border': 1})

        # 各ファイルに対する処理
        for file_idx, (file_name, df) in enumerate(files_data):
            # ファイル名を記入
            worksheet.write(f'L{file_idx * 3 + 1}', file_name, cell_format)

            # XデータとYデータを記入
            x_col = 11 + file_idx * 4  # Xデータの列（L, P, T列…）
            y_col = 12 + file_idx * 4  # Yデータの列（M, Q, U列…）
            calc_col = 13 + file_idx * 4  # 計算結果の列（N, R, V列…）

            worksheet.write(f'{chr(65 + x_col)}{file_idx * 3 + 2}', 'WL', cell_format)  # Xデータの列名
            worksheet.write(f'{chr(65 + y_col)}{file_idx * 3 + 2}', 'Abs', cell_format)  # Yデータの列名

            # Xデータ、Yデータ、計算結果を入力
            for i, (x, y) in enumerate(zip(df["X"], df["Y"])):
                try:
                    # xとyが数値であることを確認し、数値に変換
                    x = float(x) if isinstance(x, (int, float, str)) else None
                    y = float(y) if isinstance(y, (int, float, str)) else None
                except ValueError:
                    # 数値に変換できない場合はその行をスキップ
                    x, y = None, None
                
                if x is not None and y is not None:
                    worksheet.write(i + 2 + file_idx * len(df), x, cell_format)  # Xデータ
                    worksheet.write(i + 2 + file_idx * len(df), y_col, y, cell_format)  # Yデータ

            # 計算用のN1、N2、R1、R2、V1、V2などを記入
            worksheet.write(f'{chr(65 + calc_col)}{file_idx * 3 + 2}', 1, border_format)  # N1/R1/V1セルに1
            worksheet.write(f'{chr(65 + calc_col)}{file_idx * 3 + 3}', 0, border_format)  # N2/R2/V2セルに0
            worksheet.write_formula(f'{chr(65 + calc_col)}{file_idx * 3 + 4}', f"={chr(65 + y_col)}{file_idx * 3 + 4}*${chr(65 + calc_col)}${file_idx * 3 + 2}+${chr(65 + calc_col)}${file_idx * 3 + 3}", cell_format)  # 計算式を設定

            for i in range(1, len(df)):
                worksheet.write_formula(f'{chr(65 + calc_col)}{i + 2 + file_idx * len(df)}', f"={chr(65 + y_col)}{i + 3 + file_idx * len(df)}*${chr(65 + calc_col)}${file_idx * 3 + 2}+${chr(65 + calc_col)}${file_idx * 3 + 3}", cell_format)

            # 次のデータが追加される列（P, Q, R…）
            calc_col += 3  # 1列分を使用した後、次の列にデータを追加

        # グラフを作成
        chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})

        # 各ファイルのXデータと補正されたYデータをグラフに追加
        for file_idx, (file_name, df) in enumerate(files_data):
            x_col = 11 + file_idx * 4  # Xデータの列（L, P, T列…）
            y_col = 12 + file_idx * 4  # Yデータの列（M, Q, U列…）
            calc_col = 13 + file_idx * 4  # 計算結果の列（N, R, V列…）

            chart.add_series({
                'categories': f"=Data!${chr(65 + x_col)}$3:${chr(65 + x_col)}${len(df) + 2}",
                'values': f"=Data!${chr(65 + calc_col)}$3:${chr(65 + calc_col)}${len(df) + 2}",
                'marker': {'type': 'none'},
                'line': {'color': '#008EC0', 'width': 1.5},  # 線の色と太さ
            })

        # グラフのプロパティ設定
        chart.set_size({'width': 533, 'height': 377})  # 幅13.3cm, 高さ10cm
        chart.set_chartarea({'border': {'none': True}, 'fill': {'none': True}})
        chart.set_plotarea({'border': {'color': 'black', 'width': 1.5}, 'fill': {'none': True}})
        chart.set_legend({'none': True})  # 凡例を非表示

        # 横軸設定 (開始範囲を300に固定)
        chart.set_x_axis({
            'line': {'color': 'black', 'width': 1.5},
            'major_tick_mark': 'inside',
            'major_unit': 100,  # 横軸の目盛り間隔
            'min': 300,  # 横軸の開始範囲を300に固定
            'max': df['X'].max(),  # 横軸の最大値をXデータの最大値に設定
            'reverse': False,
            'name': 'Wavelength / nm',  # 横軸ラベルを設定
            'num_font': {'color': 'black', 'size': 16, 'name': 'Arial'},  # 数値のフォント設定
            'name_font': {'color': 'black', 'size': 16, 'name': 'Arial', 'bold': False},  # ラベルのフォント設定（太字解除）
        })

        # 縦軸設定 (0.1刻み)
        chart.set_y_axis({
            'line': {'color': 'black', 'width': 1.5},
            'major_tick_mark': 'inside',
            'name': 'Absorbance',  # 縦軸ラベルを設定
            'major_gridlines': {'visible': False},  # 縦軸の目盛線を削除
            'num_font': {'color': 'black', 'size': 16, 'name': 'Arial'},  # 数値のフォント設定
            'name_font': {'color': 'black', 'size': 16, 'name': 'Arial', 'bold': False},  # ラベルのフォント設定（太字解除）
            'major_unit': 0.1,  # 縦軸の目盛り間隔を0.1に設定
        })

        # グラフを配置
        worksheet.insert_chart('A3', chart)

    return output.getvalue()
