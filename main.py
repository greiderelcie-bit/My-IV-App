import flet as ft
import openpyxl

def main(page: ft.Page):
    page.title = "I-V 数据提取与分析"
    
    # 【回退语法】0.22.1 支持的字符串对齐方式
    page.vertical_alignment = "center"
    page.horizontal_alignment = "center"
    # 顶部留出 50 的空间，完美避开刘海屏
    page.padding = ft.padding.only(top=50, left=20, right=20, bottom=20)
    
    # 【回退语法】所有颜色必须为小写 colors
    page.bgcolor = ft.colors.BLUE_GREY_50 
    page.theme_mode = "light"
    page.scroll = "auto"

    # 【回退语法】字体粗细统一用小写字符串 "bold"
    txt_fitting_result = ft.Text("拟合结果 y = 待计算", size=16, weight="bold", color=ft.colors.BLUE_800)
    log_text = ft.Text(value="", size=12)
    
    chart_container = ft.Container(
        content=ft.Text("请先提取 Excel 数据以生成图表", color=ft.colors.GREY_400),
        alignment=ft.alignment.center,
        height=300,
        bgcolor=ft.colors.WHITE,
        border_radius=12,
        padding=10,
        shadow=ft.BoxShadow(spread_radius=1, blur_radius=5, color=ft.colors.GREY_300)
    )

    process_dialog = ft.AlertDialog(
        modal=True,
        title=ft.Text("操作与进程提示"),
        content=ft.Container(
            content=ft.Column([log_text], scroll="auto"),
            width=300, height=150
        ),
        # 【回退语法】加回 text= 参数
        actions=[ft.TextButton(text="关闭", on_click=lambda e: close_dialog())],
        actions_alignment="end",
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

    def process_excel_file(e: ft.FilePickerResultEvent):
        if not e.files:
            return
        
        filepath = e.files[0].path
        log_text.value = "" 
        open_dialog()
        append_log(f"1. 选择文件: {e.files[0].name}")
        
        try:
            append_log("2. 正在读取数据...")
            wb = openpyxl.load_workbook(filepath, data_only=True)
            sheet = wb.active

            v_list, i1_list, i2_list, results = [], [], [], []
            closest_v = None
            min_v_diff = float('inf')
            target_v_result = 0.0

            for row in sheet.iter_rows(values_only=True):
                if not row or len(row) < 3:
                    continue
                try:
                    v, i1, i2 = float(row[0]), float(row[1]), float(row[2])
                    
                    v_list.append(v)
                    i1_list.append(i1)
                    i2_list.append(i2)

                    x = (i1 - i2) / i1 if i1 != 0 else 0 
                    y = 8.699 + 8.03713 * x
                    results.append((v, x, y))

                    diff = abs(v - 1.0)
                    if diff < min_v_diff:
                        min_v_diff = diff
                        closest_v = v
                        target_v_result = y
                except (ValueError, TypeError):
                    continue 

            if not v_list:
                append_log("❌ 未在表格中找到有效数据！")
                page.update()
                return

            append_log("3. 数据提取与计算完成。")
            
            chart_data_i1 = [ft.LineChartDataPoint(v, i1) for v, i1 in zip(v_list, i1_list)]
            chart_data_i2 = [ft.LineChartDataPoint(v, i2) for v, i2 in zip(v_list, i2_list)]

            chart = ft.LineChart(
                data_series=[
                    ft.LineChartData(data_points=chart_data_i1, stroke_width=2, color=ft.colors.RED_400, curved=True),
                    ft.LineChartData(data_points=chart_data_i2, stroke_width=2, color=ft.colors.TEAL_400, curved=True)
                ],
                border=ft.border.all(1, ft.colors.GREY_300),
                min_x=min(v_list), max_x=max(v_list),
                min_y=min(min(i1_list), min(i2_list)), 
                max_y=max(max(i1_list), max(i2_list)),
                expand=True,
                tooltip_bgcolor=ft.colors.BLUE_GREY_800
            )

            chart_container.content = chart
            if closest_v is not None:
                txt_fitting_result.value = f"(实际 {closest_v:.2f}V) 拟合结果 y = {target_v_result:.5f}"

            append_log("4. ✅ 操作完成！请关闭此窗口查看图表。")

        except Exception as ex:
            append_log(f"❌ 发生异常: {str(ex)}")

        page.update()

    # 【回退语法】0.22.1 支持在初始化时直接绑定事件
    file_picker = ft.FilePicker(on_result=process_excel_file)
    page.overlay.append(file_picker)

    controls_panel = ft.Container(
        bgcolor=ft.colors.WHITE, padding=15, border_radius=12,
        shadow=ft.BoxShadow(spread_radius=1, blur_radius=5, color=ft.colors.GREY_300),
        content=ft.Column(
            spacing=15,
            horizontal_alignment="center",
            controls=[
                ft.Text("数据提取与拟合", size=20, weight="bold", color=ft.colors.BLUE_GREY_800),
                # 【回退语法】重新加回 text=，并且 icon 用小写的 icons
                ft.ElevatedButton(
                    text="提取数据", 
                    icon=ft.icons.FILE_UPLOAD,
                    bgcolor=ft.colors.BLUE_100, color=ft.colors.BLUE_900, height=45,
                    on_click=lambda _: file_picker.pick_files(allowed_extensions=["xlsx"])
                ),
                ft.Divider(color=ft.colors.BLUE_GREY_100),
                txt_fitting_result
            ]
        )
    )

    page.add(
        ft.Column(
            expand=True, spacing=20,
            horizontal_alignment="center",
            controls=[controls_panel, chart_container]
        )
    )

ft.app(target=main)
