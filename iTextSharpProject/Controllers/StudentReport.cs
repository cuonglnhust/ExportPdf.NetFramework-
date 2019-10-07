using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharpProject.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace iTextSharpProject.Controllers
{
    public class StudentReport
    {
        #region Declareation 
        int _totalColumn = 3;
        Document _document;
        PdfPCell _pdfCell;
        Font _fontStyle;
        PdfPTable _pdfTable = new PdfPTable(3);
        MemoryStream _memoryStream = new MemoryStream();
        List<Student> _student = new List<Student>();
        #endregion

        public byte[] PrepareReport(List<Student> students)
        {
            _student = students;

            #region
            _document = new Document(PageSize.A4, 0f, 0f, 0f, 0f);
            _document.SetPageSize(PageSize.A4);
            _document.SetMargins(20f, 20f, 20f, 20f);
            _pdfTable.WidthPercentage = 100;
            _pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;
            _fontStyle = FontFactory.GetFont("Tahoma", 8f, 1);
            PdfWriter.GetInstance(_document, _memoryStream);
            _document.Open();
            _pdfTable.SetWidths(new float[] { 20f, 150f, 100f });
            #endregion

            this.ReportHeader();
            this.ReportBody();
            _pdfTable.HeaderRows = 2;
            _document.Add(_pdfTable);
            _document.Close();
            return _memoryStream.ToArray();
        }

        public void ReportHeader()
        {
            _fontStyle = FontFactory.GetFont("Tahoma", 11f, 1);
            _pdfCell = new PdfPCell(new Phrase("Bao cao cham cong", _fontStyle));
            _pdfCell.Colspan = _totalColumn;
            _pdfCell.HorizontalAlignment = Element.ALIGN_CENTER;
            _pdfCell.Border = 0;
            _pdfCell.BackgroundColor = BaseColor.WHITE;
            _pdfCell.ExtraParagraphSpace = 0;
            _pdfTable.AddCell(_pdfCell);


            _fontStyle = FontFactory.GetFont("Tahoma", 9f, 1);
            _pdfCell = new PdfPCell(new Phrase("Studen list", _fontStyle));
            _pdfCell.Colspan = _totalColumn;
            _pdfCell.HorizontalAlignment = Element.ALIGN_CENTER;
            _pdfCell.Border = 0;
            _pdfCell.BackgroundColor = BaseColor.WHITE;
            _pdfCell.ExtraParagraphSpace = 0;
            _pdfTable.AddCell(_pdfCell);
        }

        private void ReportBody()
        {
            #region Table Header
            _fontStyle = FontFactory.GetFont("Tahoma", 8f, 1);
            _pdfCell = new PdfPCell(new Phrase("Serial Number", _fontStyle));
            _pdfCell.HorizontalAlignment = Element.ALIGN_CENTER;
            _pdfCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            _pdfCell.BackgroundColor = BaseColor.LIGHT_GRAY;
            _pdfTable.AddCell(_pdfCell);

            _pdfCell = new PdfPCell(new Phrase("Name", _fontStyle));
            _pdfCell.HorizontalAlignment = Element.ALIGN_CENTER;
            _pdfCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            _pdfCell.BackgroundColor = BaseColor.LIGHT_GRAY;
            _pdfTable.AddCell(_pdfCell);

            _pdfCell = new PdfPCell(new Phrase("Roll", _fontStyle));
            _pdfCell.HorizontalAlignment = Element.ALIGN_CENTER;
            _pdfCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            _pdfCell.BackgroundColor = BaseColor.LIGHT_GRAY;
            _pdfTable.AddCell(_pdfCell);
            _pdfTable.CompleteRow();
            #endregion

            #region Table Body
            _fontStyle = FontFactory.GetFont("Tahoma", 8f, 1);
            int SerialNumber = 1;
            foreach (Student student in _student)
            {
                _pdfCell = new PdfPCell(new Phrase(SerialNumber++.ToString(), _fontStyle));
                _pdfCell.HorizontalAlignment = Element.ALIGN_CENTER;
                _pdfCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                _pdfCell.BackgroundColor = BaseColor.WHITE;
                _pdfTable.AddCell(_pdfCell);

                _pdfCell = new PdfPCell(new Phrase(student.Name, _fontStyle));
                _pdfCell.HorizontalAlignment = Element.ALIGN_CENTER;
                _pdfCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                _pdfCell.BackgroundColor = BaseColor.WHITE;
                _pdfTable.AddCell(_pdfCell);

                _pdfCell = new PdfPCell(new Phrase(student.Roll, _fontStyle));
                _pdfCell.HorizontalAlignment = Element.ALIGN_CENTER;
                _pdfCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                _pdfCell.BackgroundColor = BaseColor.WHITE;
                _pdfTable.AddCell(_pdfCell);

                _pdfTable.CompleteRow();
            }
            #endregion
        }
    }
}