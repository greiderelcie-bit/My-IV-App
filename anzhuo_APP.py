import flet as ft
import openpyxl


def main(page: ft.Page):
    # 1. 界面基础设置 (适配手机屏幕比例与滚动)
    page.title = "I-V 数据提取与分析"
    page.window.width = 400
    page.window.height = 750
    page.bgcolor = ft.Colors.BLUE_GREY_50
    page.theme_mode = ft.ThemeMode.LIGHT
    page.scroll = ft.ScrollMode.AUTO  # 允许手机端上下滑动

    # 2. 界面组件定义
    chart_container = ft.Container(
        content=ft.Text("请先提取 Excel 数据以生成图表", color=ft.Colors.GREY_400),
        alignment=ft.alignment.center,
        height=300,  # 固定高度适配手机屏幕
        bgcolor=ft.Colors.WHITE,
        border_radius=12,
        padding=10,
        shadow=ft.BoxShadow(spread_radius=1, blur_radius=5, color=ft.Colors.GREY_300)
    )

    txt_fitting_result = ft.Text("1V 拟合结果 y = 待计算", size=16, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_800)
    log_text = ft.Text(value="", size=12)

    process_dialog = ft.AlertDialog(
        modal=True,
        title=ft.Text("操作与进程提示"),
        content=ft.Container(
            content=ft.Column([log_text], scroll=ft.ScrollMode.AUTO),
            width=300, height=150
        ),
        actions=[ft.TextButton("关闭", on_click=lambda e: close_dialog())],
        actions_alignment=ft.MainAxisAlignment.END,
    )

    def open_dialog():
        page.overlay.append(process_dialog)
        process_dialog.open = True
        page.update()

    def close_dialog():
        process_dialog.open = False
        page.update()

    def append_log(msg):
        log_text.value += f"> {msg}\n"
        page.update()

    # 3. 核心逻辑函数
    def process_excel_file(e: ft.FilePickerResultEvent):
        if not e.files:
            return

        filepath = e.files[0].path
        log_text.value = ""
        open_dialog()
        append_log(f"1. 选择文件: {e.files[0].name}")

        try:
            append_log("2. 正在通过 openpyxl 读取数据...")
            wb = openpyxl.load_workbook(filepath, data_only=True)
            sheet = wb.active

            v_list, i1_list, i2_list, results = [], [], [], []
            closest_v = None
            min_v_diff = float('inf')
            target_v_result = 0.0

            append_log("3. 提取电压、I1、I2 并过滤表头文字...")
            # 遍历行，提取前三列
            for row in sheet.iter_rows(values_only=True):
                if not row or len(row) < 3:
                    continue
                try:
                    # 尝试将前三列转换为浮点数，遇到文字表头会自动跳过
                    v, i1, i2 = float(row[0]), float(row[1]), float(row[2])

                    v_list.append(v)
                    i1_list.append(i1)
                    i2_list.append(i2)

                    # 计算公式
                    x = (i1 - i2) / i1 if i1 != 0 else 0  # 防止除以 0
                    y = 8.699 + 8.03713 * x
                    results.append((v, x, y))

                    # 寻找最接近 1V 的值
                    diff = abs(v - 1.0)
                    if diff < min_v_diff:
                        min_v_diff = diff
                        closest_v = v
                        target_v_result = y

                except (ValueError, TypeError):
                    continue  # 跳过带有文字说明的行

            if not v_list:
                append_log("❌ 未在表格中找到有效的数字数据！")
                page.update()
                return

            append_log("4. 数据计算完成。")

            print("\n" + "=" * 55)
            print("各电压下的数据与拟合结果如下：")
            for r in results:
                print(f"电压: {r[0]:>6.2f} V  =>  x = {r[1]:>8.5f},  y = {r[2]:>8.5f}")
            print("=" * 55 + "\n")
            append_log("5. 结果已输出至后台终端。")

            append_log("6. 正在生成原生交互式图表...")

            # 构建 Flet 原生折线图数据
            chart_data_i1 = [ft.LineChartDataPoint(v, i1) for v, i1 in zip(v_list, i1_list)]
            chart_data_i2 = [ft.LineChartDataPoint(v, i2) for v, i2 in zip(v_list, i2_list)]

            chart = ft.LineChart(
                data_series=[
                    ft.LineChartData(
                        data_points=chart_data_i1,
                        stroke_width=2,
                        color=ft.Colors.RED_400,
                        curved=True,  # 曲线平滑
                    ),
                    ft.LineChartData(
                        data_points=chart_data_i2,
                        stroke_width=2,
                        color=ft.Colors.TEAL_400,
                        curved=True,
                    )
                ],
                border=ft.border.all(1, ft.Colors.GREY_300),
                min_x=min(v_list), max_x=max(v_list),
                min_y=min(min(i1_list), min(i2_list)),
                max_y=max(max(i1_list), max(i2_list)),
                expand=True,
                tooltip_bgcolor=ft.Colors.BLUE_GREY_800
            )

            chart_container.content = chart

            if closest_v is not None:
                txt_fitting_result.value = f"(实际 {closest_v:.2f}V) 1V 拟合结果 y = {target_v_result:.5f}"

            append_log("7. ✅ 操作完成！可以关闭此对话框。")

        except Exception as ex:
            append_log(f"❌ 发生异常: {str(ex)}")

        page.update()

    file_picker = ft.FilePicker(on_result=process_excel_file)
    page.overlay.append(file_picker)

    # 4. 页面排版
    controls_panel = ft.Container(
        bgcolor=ft.Colors.WHITE, padding=15, border_radius=12,
        shadow=ft.BoxShadow(spread_radius=1, blur_radius=5, color=ft.Colors.GREY_300),
        content=ft.Column(
            spacing=15,
            controls=[
                ft.Text("数据提取与拟合", size=20, weight=ft.FontWeight.W_800, color=ft.Colors.BLUE_GREY_800),
                ft.ElevatedButton(
                    text="选择 Excel 提取", icon=ft.Icons.FILE_UPLOAD,
                    bgcolor=ft.Colors.BLUE_100, color=ft.Colors.BLUE_900, height=45,
                    on_click=lambda _: file_picker.pick_files(allowed_extensions=["xlsx"])
                ),
                ft.Divider(color=ft.Colors.BLUE_GREY_100),
                txt_fitting_result
            ]
        )
    )

    page.add(
        ft.Column(
            expand=True, spacing=20,
            controls=[controls_panel, chart_container]
        )
    )


ft.app(target=main)