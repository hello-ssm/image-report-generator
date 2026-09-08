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

TEXTS = {
    "zh": {
        "intro": (
            "## 图文报告生成器\n"
            "上传多张图片，为每张图片分别填写文字描述，一键生成可下载的 PDF / Excel 报告。\n"
            "手机浏览器打开这个页面同样可以直接拍照或从相册选图上传。"
        ),
        "lang_label": "界面语言 / Language",
        "title_label": "报告标题",
        "title_default": "图文报告",
        "title_placeholder": "例如：工地巡检报告",
        "files_label": "上传图片（可多选）",
        "no_files": "请先上传图片。",
        "missing_desc": "请给每一张图片都填写描述",
        "desc_prefix": "描述 - ",
        "desc_placeholder": "请输入这张图片的文字说明...",
        "format_label": "输出格式",
        "format_choices": [("PDF", "pdf"), ("Excel", "excel")],
        "generate_btn": "生成报告",
        "output_label": "生成的报告（点击下载）",
        "page_word": "第 {n} 页",
        "generated_on": "生成日期",
    },
    "en": {
        "intro": (
            "## Image Report Generator\n"
            "Upload multiple images, write a description for each, and generate a downloadable PDF / Excel report.\n"
            "Open this page on your phone's browser to take photos or pick from your album directly."
        ),
        "lang_label": "界面语言 / Language",
        "title_label": "Report Title",
        "title_default": "Image Report",
        "title_placeholder": "e.g. Site Inspection Report",
        "files_label": "Upload Images (multiple allowed)",
        "no_files": "Please upload images first.",
        "missing_desc": "Please write a description for every image",
        "desc_prefix": "Description - ",
        "desc_placeholder": "Enter a description for this image...",
        "format_label": "Output Format",
        "format_choices": [("PDF", "pdf"), ("Excel", "excel")],
        "generate_btn": "Generate Report",
        "output_label": "Generated report (click to download)",
        "page_word": "Page {n}",
        "generated_on": "Generated on",
    },
}


class ReportPDF(FPDF):
    def __init__(self, title, page_label_template="第 {n} 页"):
        super().__init__()
        self.report_title = title
        self.page_label_template = page_label_template
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
        self.cell(0, 10, self.page_label_template.format(n=self.page_no()), align="C")


def generate_document_tool(
    image_paths: list[str],
    descriptions: list[str],
    output_format: str,
    title: str = "图文报告",
    lang: str = "zh",
) -> str:
    """Generates a PDF or Excel document combining multiple images, each with its own text description.
    Args:
        image_paths: local file paths to the images to embed in the document, in order
        descriptions: text description for each image, matching image_paths by position
        output_format: desired output format, either 'pdf' or 'excel'
        title: report title shown in the PDF header / Excel title row
        lang: language for built-in labels (page numbers, "generated on"), either 'zh' or 'en'
    """
    if len(image_paths) != len(descriptions):
        raise ValueError("image_paths and descriptions must have the same length")

    t = TEXTS.get(lang, TEXTS["zh"])
    output_format = output_format.strip().lower()
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    today = datetime.date.today().isoformat()

    if output_format == "pdf":
        output_path = f"report_{timestamp}.pdf"
        pdf = ReportPDF(title, t["page_word"])
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
        ws.cell(row=2, column=1, value=f"{t['generated_on']}: {today}").font = Font(size=9, color="787878")
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


def _build_report(image_paths, descriptions, output_format, title, lang):
    t = TEXTS.get(lang, TEXTS["zh"])
    if not image_paths:
        raise gr.Error(t["no_files"])
    if any(not d or not d.strip() for d in descriptions):
        raise gr.Error(t["missing_desc"])
    return generate_document_tool(
        image_paths=image_paths,
        descriptions=[d.strip() for d in descriptions],
        output_format=output_format,
        title=title.strip() or t["title_default"],
        lang=lang,
    )


with gr.Blocks(title="图文报告生成器 / Image Report Generator") as demo:
    lang_input = gr.Radio(
        choices=[("中文", "zh"), ("English", "en")],
        value="zh",
        label=TEXTS["zh"]["lang_label"],
    )
    intro_md = gr.Markdown(TEXTS["zh"]["intro"])
    title_input = gr.Textbox(
        label=TEXTS["zh"]["title_label"],
        value=TEXTS["zh"]["title_default"],
        placeholder=TEXTS["zh"]["title_placeholder"],
    )
    files_input = gr.File(label=TEXTS["zh"]["files_label"], file_count="multiple", file_types=["image"])

    def on_lang_change(lang, current_title):
        t = TEXTS[lang]
        known_defaults = {TEXTS["zh"]["title_default"], TEXTS["en"]["title_default"]}
        new_title_value = t["title_default"] if (not current_title or current_title in known_defaults) else current_title
        return (
            gr.update(value=t["intro"]),
            gr.update(label=t["title_label"], placeholder=t["title_placeholder"], value=new_title_value),
            gr.update(label=t["files_label"]),
        )

    lang_input.change(
        fn=on_lang_change,
        inputs=[lang_input, title_input],
        outputs=[intro_md, title_input, files_input],
    )

    @gr.render(inputs=[files_input, lang_input])
    def render_form(files, lang):
        t = TEXTS.get(lang, TEXTS["zh"])
        if not files:
            gr.Markdown(t["no_files"])
            return

        description_boxes = [
            gr.Textbox(label=f"{t['desc_prefix']}{os.path.basename(f)}", placeholder=t["desc_placeholder"])
            for f in files
        ]
        format_input = gr.Radio(choices=t["format_choices"], value="pdf", label=t["format_label"])
        generate_btn = gr.Button(t["generate_btn"], variant="primary")
        output_file = gr.File(label=t["output_label"])

        def on_generate(*values):
            *descriptions, output_format, title = values
            return _build_report(files, list(descriptions), output_format, title, lang)

        generate_btn.click(
            fn=on_generate,
            inputs=description_boxes + [format_input, title_input],
            outputs=output_file,
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 7860))
    demo.launch(server_name="0.0.0.0", server_port=port)
