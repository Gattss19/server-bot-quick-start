"""

Sample bot that echoes back messages.

This is the simplest possible bot and a great place to start if you want to build your own bot.

"""

from __future__ import annotations

from typing import AsyncIterable
from docx import Document  # Для работы с .docx
from PyPDF2 import PdfReader  # Для работы с PDF
import openpyxl  # Для работы с .xlsx

import fastapi_poe as fp
from modal import App, Image, asgi_app


class EchoBot(fp.PoeBot):
    async def get_response(
        self, request: fp.QueryRequest
    ) -> AsyncIterable[fp.PartialResponse]:
        responses = []

        # Проверяем вложения
        if request.attachments:
            for attachment in request.attachments:
                if attachment.type == "text":
                    responses.append(f"Текст из вложения: {attachment.content}")
                elif attachment.type == "image":
                    responses.append(f"Изображение получено: {attachment.name}")
                elif attachment.type == "file" and attachment.name.endswith(".docx"):
                    try:
                        # Обработка .docx файла
                        doc = Document(attachment.content)
                        doc_text = "\n".join([p.text for p in doc.paragraphs])
                        responses.append(f"Текст из .docx файла: {doc_text}")
                    except Exception as e:
                        responses.append(f"Ошибка обработки .docx файла: {str(e)}")
                elif attachment.type == "file" and attachment.name.endswith(".pdf"):
                    try:
                        # Обработка PDF файла
                        pdf_reader = PdfReader(attachment.content)
                        pdf_text = "\n".join(page.extract_text() for page in pdf_reader.pages)
                        responses.append(f"Текст из PDF файла: {pdf_text}")
                    except Exception as e:
                        responses.append(f"Ошибка обработки PDF файла: {str(e)}")
                elif attachment.type == "file" and attachment.name.endswith(".xlsx"):
                    try:
                        # Обработка .xlsx файла
                        workbook = openpyxl.load_workbook(attachment.content, data_only=True)
                        sheet = workbook.active
                        rows = []
                        for row in sheet.iter_rows(values_only=True):
                            rows.append("\t".join(str(cell) if cell is not None else "" for cell in row))
                        xlsx_text = "\n".join(rows)
                        responses.append(f"Содержимое .xlsx файла:\n{xlsx_text}")
                    except Exception as e:
                        responses.append(f"Ошибка обработки .xlsx файла: {str(e)}")

        # Если нет вложений, просто отвечаем эхо-сообщением
        if not responses:
            last_message = request.query[-1].content
            responses.append(f"Вы сказали: {last_message}")

        # Возвращаем ответ
        yield fp.PartialResponse(text="\n".join(responses))

    async def get_settings(self, setting: fp.SettingsRequest) -> fp.SettingsResponse:
        # Настройки для включения поддержки вложений
        return fp.SettingsResponse(
            allow_attachments=True,  # Включить вложения
            expand_text_attachments=True,  # Автоматически извлекать текст из файлов
            enable_image_comprehension=True,  # Включить обработку изображений
            introduction_message="Привет! Отправьте текстовый файл, изображение, PDF или Excel, и я обработаю его!"
        )


# Установка зависимостей
REQUIREMENTS = ["fastapi-poe==0.0.48", "python-docx", "PyPDF2", "openpyxl"]
image = Image.debian_slim().pip_install(*REQUIREMENTS)
app = App("echobot-poe")


@app.function(image=image)
@asgi_app()
def fastapi_app():
    bot = EchoBot()
    app = fp.make_app(bot, allow_without_key=True)
    return app
