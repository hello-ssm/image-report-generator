from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font
import os
import datetime
import gradio as gr
from pillow_heif import register_heif_opener

register_heif_opener()  # allow PIL/fpdf2/openpyxl to decode iPhone HEIC/HEIF photos

CJK_FONT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts", "NotoSansCJK-Regular.ttc")


class ReportPDF(FPDF):
    def __init__(self, title):
        super().__init__()
        self.report_title = title
        self.add_font("NotoSansSC", fname=CJK_FONT_PATH, collection_font_number=2)
        self.set_font("NotoSansSC", size=12)

    def header(self):
        self.set_font("NotoSansSC", size=16)
        self.cell(0, 12, self.report_title, align="C")
        self.ln(14)
        self.set_draw_color(180, 180, 180)
        self.line(10, self.get_y(), self.w - 10, self.get_y())
        self.ln(6)
        self.set_font("NotoSansSC", size=12)

    def footer(self):
        self.set_y(-15)
        self.set_font("NotoSansSC", size=9)
        self.set_text_color(120, 120, 120)
        self.cell(0, 10, f"第 {self.page_no()} 页", align="C")


def generate_document_tool(image_paths: list[str], descriptions: list[str], output_format: str, title: str = "图文报告") -> str:
    """Generates a PDF or Excel document combining multiple images, each with its own text description.
    Args:
        image_paths: local file paths to the images to embed in the document, in order
        descriptions: text description for each image, matching image_paths by position
        output_format: desired output format, either 'pdf' or 'excel'
        title: report title shown in the PDF header / Excel title row
    """
    if len(image_paths) != len(descriptions):
        raise ValueError("image_paths and descriptions must have the same length")

    output_format = output_format.strip().lower()
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    today = datetime.date.today().isoformat()

    if output_format == "pdf":
        output_path = f"report_{timestamp}.pdf"
        pdf = ReportPDF(title)
        for image_path, description in zip(image_paths, descriptions):
            pdf.add_page()
            pdf.multi_cell(0, 10, description)
            pdf.ln(5)
            pdf.image(image_path, w=150)
        pdf.output(output_path)
        return output_path

    elif output_format in ("excel", "xlsx"):
        output_path = f"report_{timestamp}.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.column_dimensions["A"].width = 60
        ws.cell(row=1, column=1, value=title).font = Font(size=16, bold=True)
        ws.cell(row=2, column=1, value=f"生成日期: {today}").font = Font(size=9, color="787878")
        row = 4
        for image_path, description in zip(image_paths, descriptions):
            ws.cell(row=row, column=1, value=description)
            img = XLImage(image_path)
            ws.add_image(img, f"A{row + 1}")
            row += 18  # leave room for the embedded image before the next entry
        wb.save(output_path)
        return output_path

    else:
        raise ValueError("output_format must be 'pdf' or 'excel'")


def _build_report(image_paths, descriptions, output_format, title):
    if not image_paths:
        raise gr.Error("请至少上传一张图片")
    if any(not d or not d.strip() for d in descriptions):
        raise gr.Error("请给每一张图片都填写描述")
    return generate_document_tool(
        image_paths=image_paths,
        descriptions=[d.strip() for d in descriptions],
        output_format=output_format,
        title=title.strip() or "图文报告",
    )


with gr.Blocks(title="图文报告生成器") as demo:
    gr.Markdown(
        "## 图文报告生成器\n"
        "上传多张图片，为每张图片分别填写文字描述，一键生成可下载的 PDF / Excel 报告。\n"
        "手机浏览器打开这个页面同样可以直接拍照或从相册选图上传。"
    )
    title_input = gr.Textbox(label="报告标题", value="图文报告", placeholder="例如：工地巡检报告")
    files_input = gr.File(label="上传图片（可多选）", file_count="multiple", file_types=["image"])

    @gr.render(inputs=files_input)
    def render_form(files):
        if not files:
            gr.Markdown("请先上传图片。")
            return

        description_boxes = [
            gr.Textbox(label=f"描述 - {os.path.basename(f)}", placeholder="请输入这张图片的文字说明...")
            for f in files
        ]
        format_input = gr.Radio(choices=["pdf", "excel"], value="pdf", label="输出格式")
        generate_btn = gr.Button("生成报告", variant="primary")
        output_file = gr.File(label="生成的报告（点击下载）")

        def on_generate(*values):
            *descriptions, output_format, title = values
            return _build_report(files, list(descriptions), output_format, title)

        generate_btn.click(
            fn=on_generate,
            inputs=description_boxes + [format_input, title_input],
            outputs=output_file,
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 7860))
    demo.launch(server_name="0.0.0.0", server_port=port)
