import os

import pythoncom
import win32print
import win32com.client
from flask import Flask, request
from flask_cors import CORS
from werkzeug.utils import secure_filename
from werkzeug.datastructures import FileStorage
from PyPDF2 import PdfReader, PdfWriter
from pypinyin import lazy_pinyin
from PIL import Image

app = Flask(__name__)
app.config['JSON_AS_ASCII'] = False  # 支持中文
app.config['UPLOAD_FOLDER'] = os.path.abspath('static/files')
CORS(app, supports_credentials=True)


class UploadFile(object):
    def __init__(self, f: FileStorage):
        self.fileStorage = f
        self.originFilename = f.filename
        self.secureFilename = secure_filename(''.join(lazy_pinyin(f.filename)))
        self.uniqueFilename = UploadFile.make_filename_unique(self.secureFilename)
        self.extension = str(os.path.splitext(self.secureFilename)[-1]).lower()

    @staticmethod
    def make_filename_unique(filename: str) -> str:
        from uuid import uuid4
        ident = uuid4().__str__()[:8]
        return f'{ident}_{filename}'

    def save(self):
        self.fileStorage.save(os.path.join(app.config['UPLOAD_FOLDER'], self.uniqueFilename))


class Document(object):
    def __init__(self, arg: UploadFile | str):
        """
        :param arg: UploadFile Object or filename
        """
        self.filename = None
        self.extension = None
        self.absPath = None
        self._pages = None
        self._size = None
        if isinstance(arg, UploadFile):
            self.filename = arg.uniqueFilename
            self.extension = arg.extension
        elif isinstance(arg, str):
            self.filename = arg
            self.extension = str(os.path.splitext(arg)[-1]).lower()
        else:
            raise ValueError('Must be UploadFile Object or filename')

        self.absPath = os.path.join(app.config['UPLOAD_FOLDER'], self.filename)

    @property
    def size(self) -> str:
        if self._size is None:
            st_size = os.stat(self.absPath).st_size
            if st_size < 1048576:
                self._size = str(round(st_size / 1024, 2)) + 'KB'
            else:
                self._size = str(round(st_size / 1048576, 2)) + 'MB'
        return self._size

    def convert2pdf(self) -> str:
        in_path = self.absPath
        out_path = os.path.splitext(self.absPath)[0] + '.pdf'

        if self.extension == '.pdf':
            return self.filename
        if self.extension in ['.doc', '.docx']:
            # Initialize
            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx('Word.Application')
            doc = word.Documents.Open(in_path)
            doc.SaveAs(out_path, FileFormat=17)
            doc.Close()
            word.Quit()
            # Uninitialize
            pythoncom.CoUninitialize()
        if self.extension in ['.jpg']:
            image_1 = Image.open(in_path)
            im_1 = image_1.convert('RGB')
            im_1.save(out_path)

        return out_path


class PDFDocument(Document):
    def __init__(self, arg: UploadFile | str):
        super().__init__(arg)

    @property
    def pages(self) -> int:
        if self._pages is None:
            reader = PdfReader(self.absPath)
            self._pages = len(reader.pages)
        return self._pages

    def add_to_printer(self, options: dict = None):
        args = PrinterUtil.parse_options(options)
        cmd = os.path.abspath('SumatraPDF.exe') + ' ' + args + ' ' + self.absPath
        print(cmd)
        os.system(cmd)
        return PrinterUtil.get_job_id_by_document(PrinterUtil.get_default_printer(), self.filename)

    def strict_add_to_printer(self, options: dict = None):
        lst = options['pages'].split(',')
        pages = []
        for r in lst:
            temp = r.split('-')
            if len(temp) == 2:
                pages.extend([n for n in range(int(temp[0]), int(temp[1]) + 1)])
            elif len(temp) == 1:
                pages.append(int(temp[0]))
        pages = set(pages)
        reader = PdfReader(self.absPath)
        writer = PdfWriter()
        for index in pages:
            writer.add_page(reader.pages[index - 1])
        if options['side'] == 'duplex' and len(pages) & 1:
            writer.add_blank_page()

        new_filename = os.path.splitext(self.filename)[0] + '_output' + self.extension
        output = os.path.join(app.config['UPLOAD_FOLDER'], new_filename)
        with open(output, 'wb') as f:
            writer.write(f)
        del options['pages']
        args = PrinterUtil.parse_options(options)
        cmd = os.path.abspath('SumatraPDF.exe') + ' ' + args + ' ' + output
        os.system(cmd)
        return PrinterUtil.get_job_id_by_document(PrinterUtil.get_default_printer(), new_filename)


class PrinterUtil(object):
    @staticmethod
    def trans_job_info(job_info):
        ret = {
            'jobId': job_info['JobId'],
            'printerName': job_info['pPrinterName'],
            'document': job_info['pDocument'],
            'status': job_info['Status'],
            'priority': job_info['Priority'],
            'position': job_info['Position'],
            'totalPages': job_info['TotalPages'],
            'pagesPrinted': job_info['PagesPrinted']
        }
        return ret

    @staticmethod
    def parse_options(options):
        args = ''
        if options is None:
            args += '-print-to-default -silent '
        else:
            if 'printer' in options:
                args += '-print-to ' + options['printer'] + ' '
            else:
                args += '-print-to-default '
            args += '-silent '
            keys = list(options.keys())
            if 'printer' in keys:
                keys.remove('printer')
            if len(keys) > 0:
                args += '-print-settings '
            if 'pages' in options:
                args += options['pages'] + ','
            if 'monochrome' in options and options['monochrome']:
                args += 'monochrome,'
            if 'side' in options:
                args += options['side'] + ','
            if 'paperSize' in options:
                args += 'paper=' + options['paperSize'] + ','
            if 'copies' in options:
                args += str(options['copies']) + 'x,'
        return args[:-1]

    @staticmethod
    def get_default_printer() -> str:
        return win32print.GetDefaultPrinter()

    @staticmethod
    def set_default_printer(printer: str):
        win32print.SetDefaultPrinter(printer)

    @staticmethod
    def enum_jobs(printer: str):
        handle = win32print.OpenPrinter(printer)
        jobs_info = win32print.EnumJobs(handle, 0, -1)
        win32print.ClosePrinter(handle)
        data = []
        for job in jobs_info:
            data.append(PrinterUtil.trans_job_info(job))
        return data

    @staticmethod
    def get_job(printer: str, job_id: int):
        handle = win32print.OpenPrinter(printer)
        job_info = win32print.GetJob(handle, job_id)
        win32print.ClosePrinter(handle)
        return PrinterUtil.trans_job_info(job_info)

    @staticmethod
    def get_job_id_by_document(printer: str, document: str) -> int:
        jobs = PrinterUtil.enum_jobs(printer)
        job_id = None
        for job in jobs:
            if os.path.basename(job['document']) == document:
                job_id = job['jobId']
                break
        return job_id


@app.route('/uploader', methods=['GET', 'POST'])
def uploader():
    if request.method == 'POST':
        f = request.files['file']
        upload_file = UploadFile(f)
        upload_file.save()
        doc = Document(upload_file)
        pdf = PDFDocument(doc.convert2pdf())
        return {
            'id': pdf.filename.split('_')[0],
            'originFilename': upload_file.originFilename,
            'newFilename': pdf.filename,
            'size': pdf.size,
            'pages': pdf.pages,
            'extension': upload_file.extension
        }
    else:
        return 'method not allowed'


@app.route('/get_default_printer')
def get_default_printer():
    return PrinterUtil.get_default_printer()


@app.route('/enum_jobs')
def enum_jobs():
    data = PrinterUtil.enum_jobs(PrinterUtil.get_default_printer())
    return {
        'data': data
    }


@app.route('/get_job')
def get_job():
    job_id = int(request.args.get('jobID'))
    try:
        ret = PrinterUtil.get_job(PrinterUtil.get_default_printer(), job_id)
    except Exception:
        ret = {
            "status": -5
        }
    return ret


@app.route('/print', methods=['post'])
def print_document():
    data = request.get_json()
    pdf = PDFDocument(data['filename'])
    if data['options']['side'] == 'simplex':
        job_id = pdf.add_to_printer(data['options'])
    else:
        job_id = pdf.strict_add_to_printer(data['options'])
    return {
        'jobID': job_id
    }


if __name__ == '__main__':
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    app.run()
