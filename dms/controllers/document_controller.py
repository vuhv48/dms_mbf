from odoo import http
from odoo.http import request
# import aspose.words as aw
import base64
import tempfile
import subprocess

class DocumentController(http.Controller):

    # @http.route('/web/content/doc_preview', type='http', auth='public')
    # def doc_preview(self, id):
    #     attachment = request.env['dms.file'].sudo().search([('id', '=', id)], limit=1)

    #     if attachment:
    #         if attachment.mimetype in [
    #             'application/msword',
    #             'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    #         ]:
    #             # Tạo file tạm thời để lưu DOCX
    #             with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_docx:
    #                 temp_docx.write(base64.b64decode(attachment.content))
    #                 temp_docx_path = temp_docx.name

    #             # Đường dẫn tới file PDF đầu ra
    #             pdf_file_path = tempfile.mktemp(suffix=".pdf")

    #             # Mở file DOCX và chuyển đổi sang PDF
    #             doc = aw.Document(temp_docx_path)
    #             doc.save(pdf_file_path)

    #             with open(pdf_file_path, 'rb') as pdf_file:
    #                 pdf_content = pdf_file.read()

    #             # Trả về file PDF cho người dùng
    #             response = request.make_response(pdf_content)
    #             response.headers['Content-Type'] = 'application/pdf'
    #             response.headers['Content-Disposition'] = f'inline; filename="{attachment.name}.pdf"'
    #             return response

    #     return request.not_found()
    
    
    @http.route('/web/content/ppt_preview', type='http', auth='public')
    def ppt_preview(self, id):
        attachment = request.env['dms.file'].sudo().search([('id', '=', id)], limit=1)

        if attachment:
            if attachment.mimetype in [
                'application/vnd.ms-powerpoint',
                'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                'application/zip'
            ]:
                # Tạo tệp tạm thời để lưu PowerPoint
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as temp_pptx:
                    temp_pptx.write(base64.b64decode(attachment.content))
                    temp_pptx_path = temp_pptx.name

                # Đường dẫn tệp PDF đầu ra
                pdf_file_path = tempfile.mktemp(suffix='.pdf')

                # Chuyển đổi PowerPoint sang PDF
                try:
                    subprocess.run(['unoconv', '-f', 'pdf', '-o', pdf_file_path, temp_pptx_path], check=True)

                    # Đọc nội dung PDF và trả về phản hồi
                    with open(pdf_file_path, 'rb') as pdf_file:
                        pdf_content = pdf_file.read()

                    response = request.make_response(pdf_content)
                    response.headers['Content-Type'] = 'application/pdf'
                    response.headers['Content-Disposition'] = f'inline; filename="{attachment.name.replace(".pptx", ".pdf")}"'
                    return response

                except subprocess.CalledProcessError as e:
                    return request.not_found()


        return request.not_found()
    
    
    @http.route('/web/content/docx_preview', type='http', auth='public')
    def docx_preview(self, id):
        attachment = request.env['dms.file'].sudo().search([('id', '=', id)], limit=1)

        if attachment:
            if attachment.mimetype in [
                'application/msword',
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            ]:
                # Tạo file tạm thời để lưu DOCX
                with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_docx:
                    temp_docx.write(base64.b64decode(attachment.content))
                    temp_docx_path = temp_docx.name

                # Đường dẫn tới file PDF đầu ra
                pdf_file_path = tempfile.mktemp(suffix=".pdf")

                # Chuyển đổi DOCX sang PDF bằng unoconv
                try:
                    subprocess.run(['unoconv', '-f', 'pdf', '-o', pdf_file_path, temp_docx_path], check=True)

                    # Đọc nội dung PDF để trả về
                    with open(pdf_file_path, 'rb') as pdf_file:
                        pdf_content = pdf_file.read()

                    # Trả về file PDF cho người dùng
                    response = request.make_response(pdf_content)
                    response.headers['Content-Type'] = 'application/pdf'
                    response.headers['Content-Disposition'] = f'inline; filename="{attachment.name}.pdf"'
                    return response
                except subprocess.CalledProcessError as e:
                    # Xử lý lỗi nếu có
                    return request.not_found()  # Hoặc trả về một thông báo lỗi cụ thể


        return request.not_found()