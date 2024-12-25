# my view code 


from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, Table, TableStyle
from io import BytesIO
import os
import pandas as pd
from django.conf import settings
from django.http import HttpResponse
from reportlab.pdfgen import canvas
from .forms import PDFForm
from django.shortcuts import render

def generate_pdf_view(request):
    if request.method == 'POST':
        form = PDFForm(request.POST, request.FILES)
        if form.is_valid():
            pdf_data = form.save()
            buffer = BytesIO()
            p = canvas.Canvas(buffer, pagesize=A4)
            width, height = A4
            styles = getSampleStyleSheet()

            if form.cleaned_data.get('include_warrantyletter'):
            # Add letterhead and signature images if available
                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 40, height - 100, width=600, height=100)

                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 40, height - 500, width=300, height=40)

                if pdf_data.excel:
                    excel_path = os.path.join(settings.MEDIA_ROOT, pdf_data.excel.name)

                # Add other form data to the PDF
                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    p.drawString(400, height - 160, "Date: " + date)

                # Self declaration letter page
                elements = []

                # Creating the paragraphs
                address = f"""
                    To,<br/>
                    {pdf_data.designation or ""}<br/>
                    {pdf_data.institution_name or ""}<br/>
                    {pdf_data.address or ""}
                """
                elements.append(Paragraph(address, styles['BodyText']))

                elements.append(Paragraph("<b>Self declaration letter</b>", styles['Title']))

                body_text = """
                    Dear Sir/madam,<br/>
                    This refers to you that related to above subject with respect to the procurement procedure of 
                    the organization. We Fine Surgicals Nepal Pvt. Ltd. declare that we are eligible to participate 
                    in the bidding process. We have no conflict of interest with Bidder with respect to the proposed 
                    bid procurement proceedings and have no pending Litigation with the bidder.<br/><br/>
                    Thank you sincerely yours,<br/>
                    {0}<br/>
                    {1}
                """.format(pdf_data.proprietor_name or "", pdf_data.designation or "")
                elements.append(Paragraph(body_text, styles['BodyText']))

                # Render elements for the Self Declaration letter
                for element in elements:
                    w, h = element.wrap(width - 100, height - 320)  # Set width and adjust height
                    element.drawOn(p, 50, height - 200 - h)
                    height -= h + 20  # Adjust height for the next paragraph

                p.showPage()  # Move to the next page for the next letter
            if form.cleaned_data.get('include_commitmentletter'):
            # Warranty commitment letter page
                

                elements = []

                # Creating the paragraphs for warranty commitment
                elements.append(Paragraph(address, styles['BodyText']))

                elements.append(Paragraph("<b>warranty/commitment letter</b>", styles['Title']))

                commitment_text = """
                    Dear Sir/madam,<br/>
                    This refers to your Tender for The Procurement of Turn-key establishment (infrastructure 
                    modification). We, Fine Surgicals Nepal Pvt. Ltd., declare that we are eligible to participate 
                    in the bidding process. We have no conflict of interest with Bidder with respect to the proposed 
                    bid procurement proceedings and have no pending Litigation with the bidder.<br/><br/>
                    Thank you sincerely yours,<br/>
                    {0}<br/>
                    {1}
                """.format(pdf_data.proprietor_name or "", pdf_data.designation or "")
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                # Render elements for the warranty commitment letter
                for element in elements:
                    w, h = element.wrap(width - 100, height - 320)
                    element.drawOn(p, 50, height - 200 - h)
                    height -= h + 20

                p.showPage()
            if form.cleaned_data.get('include_table_data'):
            # BOQ page - rendering the Excel data
                data = pd.read_excel(excel_path, engine='openpyxl')
                data_list = [data.columns.tolist()] + data.values.tolist()

                # Prepare the table data with Paragraphs
                table_data = []
                for row in data_list:
                    wrapped_row = [Paragraph(str(cell), styles["BodyText"]) for cell in row]
                    table_data.append(wrapped_row)

                # Column widths
                available_width = 500
                column_widths = [available_width / len(data.columns) for _ in data.columns]

                # Create and style the table
                table = Table(table_data, colWidths=column_widths)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))

                # Draw the table
                table.wrapOn(p, width, height)
                table.drawOn(p, 50, height - 600)

            p.save()

            # Return the PDF as a response
            buffer.seek(0)
            return HttpResponse(buffer, content_type='application/pdf', headers={'Content-Disposition': 'attachment; filename="generated_document.pdf"'})

    else:
        form = PDFForm()

    return render(request, 'pdfgenerator/forms.html', {'form': form})






from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, Table, TableStyle,Frame
from io import BytesIO
import os
import pandas as pd
from django.conf import settings
from django.http import HttpResponse
from reportlab.pdfgen import canvas
from .forms import PDFForm
from django.shortcuts import render


def rupee_format_get_d(x_str):
    arr_1 = ["One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", ""]
    return arr_1[int(x_str) - 1] if int(x_str) > 0 else ""

def rupee_format_get_t(x_str):
    arr1 = ["Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
    arr2 = ["", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
    result = ""

    if int(x_str[0]) == 1:
        result = arr1[int(x_str[1])]
    else:
        if int(x_str[0]) > 0:
            result = arr2[int(x_str[0]) - 1]
        result += rupee_format_get_d(x_str[1])

    return result

def rupee_format_get_h(x_str_h, x_lp):
    x_str_h = x_str_h.zfill(3)  # Pad to ensure it's 3 digits
    result = ""

    if int(x_str_h) < 1:
        return ""

    if x_str_h[0] != "0":
        if x_lp > 0:
            result = rupee_format_get_d(x_str_h[0]) + " Lakh "
        else:
            result = rupee_format_get_d(x_str_h[0]) + " Hundred "

    if x_str_h[1] != "0":
        result += rupee_format_get_t(x_str_h[1:])
    else:
        result += rupee_format_get_d(x_str_h[2])

    return result

def rupee_format(s_num):
    arr_place = ["", "", " Thousand ", " Lakhs ", " Crores ", " Trillion ", "", "", "", ""]
    if s_num == "":
        return ""

    x_num_str = str(s_num).strip()

    if x_num_str == "":
        return ""

    if float(x_num_str) > 999999999.99:
        return "Digit exceeds Maximum limit"

    x_dp_int = x_num_str.find(".")

    x_rstr_paisas = ""  # Initialize x_rstr_paisas here

    if x_dp_int > 0:
        if len(x_num_str) - x_dp_int == 1:
            x_rstr_paisas = rupee_format_get_t(x_num_str[x_dp_int + 1:x_dp_int + 2] + "0")
        elif len(x_num_str) - x_dp_int > 1:
            x_rstr_paisas = rupee_format_get_t(x_num_str[x_dp_int + 1:x_dp_int + 3])

        x_num_str = x_num_str[:x_dp_int]

    x_f = 1
    x_rstr = ""
    x_lp = 0

    while x_num_str != "":
        if x_f >= 2:
            x_temp = x_num_str[-2:]
        else:
            if len(x_num_str) == 2:
                x_temp = x_num_str[-2:]
            elif len(x_num_str) == 1:
                x_temp = x_num_str[-1:]
            else:
                x_temp = x_num_str[-3:]

        x_str_temp = ""

        if int(x_temp) > 99:
            x_str_temp = rupee_format_get_h(x_temp, x_lp)
            if "Lac" not in x_str_temp:
                x_lp += 1
        elif int(x_temp) <= 99 and int(x_temp) > 9:
            x_str_temp = rupee_format_get_t(x_temp)
        elif int(x_temp) < 10:
            x_str_temp = rupee_format_get_d(x_temp)

        if x_str_temp != "":
            x_rstr = x_str_temp + arr_place[x_f] + x_rstr

        if x_f == 2:
            if len(x_num_str) == 1:
                x_num_str = ""
            else:
                x_num_str = x_num_str[:-2]
        elif x_f == 3:
            if len(x_num_str) >= 3:
                x_num_str = x_num_str[:-2]
            else:
                x_num_str = ""
        elif x_f == 4:
            x_num_str = ""
        else:
            if len(x_num_str) <= 2:
                x_num_str = ""
            else:
                x_num_str = x_num_str[:-3]

        x_f += 1

    if x_rstr == "":
        x_rstr = "No Rupees"
    else:
        x_rstr = "Rupees " + x_rstr

    if x_rstr_paisas != "":
        x_rstr_paisas = " and " + x_rstr_paisas + " Paisas"

    return x_rstr + x_rstr_paisas + " Only"



def generate_pdf_view(request):
    
    if request.method == 'POST':
        form = PDFForm(request.POST, request.FILES)
        if form.is_valid():
            pdf_data = form.save()
            buffer = BytesIO()
            p = canvas.Canvas(buffer, pagesize=A4)
            width, height = A4
            styles = getSampleStyleSheet()
            styles['BodyText'].fontSize = 12
            styles['BodyText'].leading = 20
            # words = rupee_format(pdf_data.amount)

            if form.cleaned_data.get('include_selfdeclarationletter'):

                # Draw letterhead image at the top, if available
                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 40, height - 100, width=600, height=100)

                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    p.drawString(400, height - 160, "Date: " + date)

                # Add the address and content text
                elements = []

                # Address block as a Paragraph
                address = f"""
                    To,<br/>
                    {pdf_data.designation or ""}<br/>
                    {pdf_data.institution_name or ""}<br/>
                    {pdf_data.address or ""}
                """
                elements.append(Paragraph(address, styles['BodyText']))

                # Declaration Title
                elements.append(Paragraph("<b>Self Declaration Letter</b>", styles['Title']))

                # Body Content
                body_text = """
                    Dear Sir/madam,<br/>
                    This refers to you that related to above subject with respect to the procurement procedure of 
                    the organization. We Fine Surgicals Nepal Pvt. Ltd. declare that we are eligible to participate 
                    in the bidding process.<br/> We have no conflict of interest with Bidder with respect to the proposed 
                    bid procurement proceedings and have no pending Litigation with the bidder.<br/><br/>
                    Thank you sincerely yours,
                """
                elements.append(Paragraph(body_text, styles['BodyText']))

                # Draw elements in order
                y_position = height - 200  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=40)
                    y_position -= 100  # Adjust y position below the signature image

                # Add proprietor_name and designation below the signature image
                # proprietor_info = f"""
                #     {pdf_data.proprietor_name or ""}<br/>
                #     {pdf_data.designation or ""}
                # """
                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.designation)

                p.showPage()  # Move to the next page if needed

            if form.cleaned_data.get('include_warrantyletter'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 40, height - 100, width=600, height=100)

                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    p.drawString(400, height - 160, "Date: " + date)

            # Warranty commitment letter page
                # Address block as a Paragraph
                address = f"""
                    To,<br/>
                    {pdf_data.designation or ""}<br/>
                    {pdf_data.institution_name or ""}<br/>
                    {pdf_data.address or ""}
                """

                elements = []

                # Creating the paragraphs for warranty commitment
                elements.append(Paragraph(address, styles['BodyText']))

                elements.append(Paragraph("<b>warranty/commitment letter</b>", styles['Title']))

                commitment_text = f"""
                    Dear Sir/madam,<br/>
                    This refers to your Tender for The Procurement of {pdf_data.subject}. We, Fine Surgicals Nepal Pvt. Ltd., declare that we are eligible to participate 
                    in the bidding process.We Fine Surgicals Nepal Pvt. Ltd ., if awarded for the tender we warrant that the Goods are new, unused &amp; of
                    the most recent or current models &amp; that they incorporate all recent improvements in design &amp; materials
                    unless provided otherwise in the Contract.<br/><br/>
                    We further warrant that the Goods shall be free from defects arising from any act or commission of us or
                    arising from design, materials &amp; workmanship under normal use in the conditions prevailing in the country
                    of final destination.<br/><br/>
                    The comprehensive warranty &amp; service warranty for offered product shall remain valid for the required
                    periods after the Goods or any portion thereof as the case may be, have been delivered to &amp; accepted at the
                    final destination.<br/><br/>
                    We further ensured that during the warranty period we will provide Preventive Maintenance (PPM) along
                    with corrective / breakdown maintenance whenever required.<br/><br/>
                    Thank you sincerely yours,<br/>
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 200  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=40)
                    y_position -= 100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.designation)
                p.showPage()
            

            if form.cleaned_data.get('include_manufactureletter'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 40, height - 100, width=600, height=100)

                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    p.drawString(400, height - 160, "Date: " + date)

            # Warranty commitment letter page
                # Address block as a Paragraph
                address = f"""
                    To,<br/>
                    {pdf_data.designation or ""}<br/>
                    {pdf_data.institution_name or ""}<br/>
                    {pdf_data.address or ""}
                """

                elements = []

                # Creating the paragraphs for warranty commitment
                elements.append(Paragraph(address, styles['BodyText']))

                elements.append(Paragraph("<b>Date of Manufacture and brand new machine</b>", styles['Title']))

                commitment_text = f"""
                    Dear Sir/madam,<br/>
                    This refers to your Tender for <b> The Procurement of {pdf_data.subject}</b>. We, Fine Surgicals Nepal Pvt. Ltd., declare that the offered equipment shall
                    be brand new manufactured after placement of the order. <br/><br/>             
                    We shall also submit certification with the date of manufacture of the machine along with shipment
                    of the system..<br/><br/>
                    Thank you sincerely yours,<br/>
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 200  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=40)
                    y_position -= 100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.designation)
                p.showPage()


            if form.cleaned_data.get('include_deliverycommitmentletter'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 40, height - 100, width=600, height=100)

                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    p.drawString(400, height - 160, "Date: " + date)

            # Warranty commitment letter page
                # Address block as a Paragraph
                address = f"""
                    To,<br/>
                    {pdf_data.designation or ""}<br/>
                    {pdf_data.institution_name or ""}<br/>
                    {pdf_data.address or ""}
                """

                elements = []

                # Creating the paragraphs for warranty commitment
                elements.append(Paragraph(address, styles['BodyText']))

                elements.append(Paragraph("<b>Delivery commitment letter</b>", styles['Title']))

                commitment_text = f"""
                    Dear Sir/madam,<br/>
                    This refers to your Tender for <b> The Procurement of {pdf_data.subject}</b>. We, Fine Surgicals Nepal Pvt. Ltd., declare that We Fine Surgicals Nepal Pvt.Ltd., if we are awarded the
                    contract, the Machine can be delivered as per the timeline of the contract after conformation.<br/><br/>
                    Thank you sincerely yours,<br/>
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 200  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=40)
                    y_position -= 100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.designation)
                p.showPage()


            # installation 

            if form.cleaned_data.get('include_installationletter'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 40, height - 100, width=600, height=100)

                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    p.drawString(400, height - 160, "Date: " + date)

            # Warranty commitment letter page
                # Address block as a Paragraph
                address = f"""
                    To,<br/>
                    {pdf_data.designation or ""}<br/>
                    {pdf_data.institution_name or ""}<br/>
                    {pdf_data.address or ""}
                """

                elements = []

                # Creating the paragraphs for warranty commitment
                elements.append(Paragraph(address, styles['BodyText']))

                elements.append(Paragraph("<b>Installation Demonstration And Training </b>", styles['Title']))

                commitment_text = f"""
                    Dear Sir/madam,<br/>
                    This refers to your Tender for <b> The Procurement of {pdf_data.subject}</b>. We, Fine Surgicals Nepal Pvt. Ltd. If we are awarded the contract installation &amp;
                     demonstration of the equipment shall be carried out by the trained engineers.<br/>.They shall also impart training to the engineer, technicians &amp; users at site during Installation of the system.
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 200  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=40)
                    y_position -= 100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.designation)
                p.showPage()

            if form.cleaned_data.get('include_attorneyletter'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 40, height - 100, width=600, height=100)

                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    p.drawString(400, height - 160, "Date: " + date)

            # Warranty commitment letter page
                # Address block as a Paragraph
                address = f"""
                    To,<br/>
                    {pdf_data.designation or ""}<br/>
                    {pdf_data.institution_name or ""}<br/>
                    {pdf_data.address or ""}
                """

                elements = []

                # Creating the paragraphs for warranty commitment
                elements.append(Paragraph(address, styles['BodyText']))

                elements.append(Paragraph("<b>Power of attorney </b>", styles['Title']))

                commitment_text = f"""
                    Dear Sir/madam,<br/>
                    We, Fine Surgicals Nepal Pvt. Ltd. Here by to assign{pdf_data.proprietor_name} {pdf_data.proprietor_name} of this company to handle all tender and quotation related to <b> The Procurement of {pdf_data.subject}</b>.<br/> he is authorized to do all necessary documentation on our behalf and purpose 
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 200  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=40)
                    y_position -= 100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.designation)
                p.showPage()

            if form.cleaned_data.get('include_bidsubmissionform'):
                
                words = rupee_format(int(pdf_data.amount))
                print("Formatted Words:", words)

                y_position = height - 200 
                
                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 40, height - 100, width=600, height=100)

                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    p.drawString(400, height - 160, "Date: " + date)

            # Warranty commitment letter page
                # Address block as a Paragraph
                address = f"""
                    To,<br/>
                    {pdf_data.designation or ""}<br/>
                    {pdf_data.institution_name or ""}<br/>
                    {pdf_data.address or ""}
                """

                elements = []

                # Creating the paragraphs for warranty commitment
                elements.append(Paragraph(address, styles['BodyText']))

                elements.append(Paragraph("<b>Bidsubmission forms </b>", styles['Title']))

                commitment_text = f"""
                    we,the undersigned declare that,<br/>
                    (a) We have examined &amp; have no reservations to the Bidding Documents, including Addenda No.: No addenda.<br/><br/>
                    (b) We offer to supply in conformity with the Bidding Documents &amp; in accordance with the Delivery Schedules
                    specified in the Schedule of Requirements the following Goods &amp; related Services:{pdf_data.subject} . The total
                    price of our Bid, excluding any discounts offered in item (d) below, is; Nrs {words} (Nrs. only)
                    excluding Tax and Vat.<br/><br/>
                    (c) The discounts offered and the methodology for their application are: Not applicable.<br/><br/>
                    (d) Our bid shall be valid for the period of 90 days from the date fixed for the bid submission deadline in
                    accordance with the Bidding Document, &amp; it shall remain binding upon us &amp; may be accepted at any time
                    before the expiration of that period<br/><br/>
                    (e) If our Bid is accepted, we commit to obtain a Performance Security in the amount as specified in 1TB 41 for
                    the due performance of the Contract<br/><br/>
                    (f) We are not participating, as Bidders, in more than one Bid in this bidding process, other than alternative
                    offers in accordance with the Biding Document.<br/><br/>
                    
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))
                
                
                y_position = height - 200  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20
                
                
                p.showPage()

            if form.cleaned_data.get('include_bidsubmissionform'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 40, height - 100, width=600, height=100)



                elements = []


                commitment_text = f"""
                    (g) The following commissions, gratuities, or fees, if any, have been paid or are to be paid with respect to the
                    bidding process or execution of the Contract:<br/><br/>

                   <b> Name of Receipt &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; Address &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; Reason &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;

                    Amount<br/><br/><br/>

                    &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; NONE<br/><br/></b>

                    (h) We understand that this bid, together with your written acceptance thereof included in your notification of
                    Ward shall constitute a binding contract between us, until a formal contract is prepared &amp; executed.<br/><br/>
                    (i) We understand that you are not bound to accept the lowest evaluated bid or any other bid that you may
                    receive. <br/><br/>
                    (j) We declare that we are have not been blacklisted as per 1TB 3.4 and no conflict of interest in the proposed
                    procurement proceedings &amp; we have not been punished for an offence relating to the concerned profession
                    of business.<br/><br/>
                    (k) We agree to permit GoN/DP or its representative to inspect our accounts &amp; records &amp; other documents
                    relating to the bid submission &amp; to have them audited by auditors appointed by the GoN/DP.<br/><br/>
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 150  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=40)
                    y_position -= 100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.designation)
                p.showPage()

# letter of technicxal bid

            if form.cleaned_data.get('include_technicalbid'):

                y_position = height - 200 
                
                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 40, height - 100, width=600, height=100)

                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    p.drawString(400, height - 160, "Date: " + date)

            # Warranty commitment letter page
                # Address block as a Paragraph
                address = f"""
                    To,<br/>
                    {pdf_data.designation or ""}<br/>
                    {pdf_data.institution_name or ""}<br/>
                    {pdf_data.address or ""}
                """

                elements = []

                # Creating the paragraphs for warranty commitment
                elements.append(Paragraph(address, styles['BodyText']))

                elements.append(Paragraph("<b>letter of technical bid </b>", styles['Title']))

                commitment_text = f"""
                    (a) We have examined and have no reservations to the Bidding Document, including Addenda issued in
                    accordance with Instructions to Bidders (ITB) Clause 9;<br/><br/>
                    (b) We offer to supply in conformity with the Bidding Document and in accordance with the delivery
                    schedule specified in the Section V (Schedule of Requirements), the following Goods and Related
                    Services: The Procurement of {pdf_data.subject} .<br/><br/>
                    (c) Our Bid consisting of the Technical Bid and the Price Bid shall be valid for a period of 90 days from the
                    date fixed for the bid submission deadline in accordance with the Bidding Document, and it shall
                    remain binding upon us and may be accepted at any time before the expiration of that period;<br/><br/>
                    (d) Our firm, including any subcontractors or suppliers for any part of the Contract, has nationalities from
                    eligible countries in accordance with ITB 4.8 and meets the requirements of ITB 3.4 &amp; 3.5;<br/><br/>
                    (e) We are not participating, as a Bidder or as a subcontractor/supplier, in more than one Bid in this bidding
                    process in accordance with ITB 4.3(e), other than alternative Bids in accordance with ITB 14;<br/><br/>
                    (f) Our firm, its affiliates or subsidiaries, including any Subcontractors or Suppliers for any part of the
                    contract, has not been declared ineligible by DP, under the Purchaser’s country laws or official
                    regulations or by an act of compliance with a decision of the United Nations Security Council;<br/><br/>
                    
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))
                
                
                y_position = height - 200  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20
                
                
                p.showPage()

            if form.cleaned_data.get('include_technicalbid'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 40, height - 100, width=600, height=100)



                elements = []


                commitment_text = f"""
                    (g) We are not a government owned entity/we are a government owned entity but meet the requirements of
                    ITB 4.5;<br/><br/>
                    (h) We declare that, we including any subcontractors or suppliers for any part of the contract do not have
                    any conflict of interest in accordance with ITB 4.3 and we have not been punished for an offense
                    relating to the concerned profession or business.<br/><br/>

                    (i) The following commissions, gratuities, or fees, if any, have been paid or are to be paid with respect to the
                    bidding process or execution of the Contract:<br/><br/>

                   <b> Name of Receipt &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; Address &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; Reason &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;

                    Amount<br/><br/><br/>

                    &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; NONE<br/><br/></b>

                    (j) We declare that we are solely responsible for the authenticity of the documents submitted by us. The
                        document and information submitted by us are true and correct. If any document/information given is
                        found to be concealed at a later date, we shall accept any legal actions by the purchaser.<br/><br/>
                    (k) We agree to permit GoN/DP or its representative to inspect our accounts and records and other
                        documents relating to the bid submission and to have them audited by auditors appointed by the
                        GoN/DP.<br/><br/>
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 150  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=40)
                    y_position -= 100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.designation)
                p.showPage()



            p.save()

            # Return the PDF as a response
            buffer.seek(0)
            return HttpResponse(buffer, content_type='application/pdf', headers={'Content-Disposition': 'attachment; filename="generated_document.pdf"'})

    else:
        form = PDFForm()

    return render(request, 'pdfgenerator/forms.html', {'form': form})




# code of table thats working 


from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, Table, TableStyle,Frame
from io import BytesIO
import os
import pandas as pd
from django.conf import settings
from django.http import HttpResponse
from reportlab.pdfgen import canvas
from .forms import PDFForm
from django.shortcuts import render

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

def generate_pdf_view(request):
        
    if request.method == 'POST':
        form = PDFForm(request.POST, request.FILES)
        if form.is_valid():
            pdf_data = form.save()
            buffer = BytesIO()
            p = canvas.Canvas(buffer, pagesize=A4)
            width, height = A4
            styles = getSampleStyleSheet()
            styles['BodyText'].fontSize = 12
            styles['BodyText'].leading = 20
            # words = rupee_format(pdf_data.amount)

            if form.cleaned_data.get('include_table_data'):
                
                # BOQ page - rendering the Excel data
                if pdf_data.excel:
                    excel_path = os.path.join(settings.MEDIA_ROOT, pdf_data.excel.name)

                data = pd.read_excel(excel_path, engine='openpyxl')
                data_list = [data.columns.tolist()] + data.values.tolist()

                # Prepare the table data with Paragraphs
                table_data = []
                for row in data_list:
                    wrapped_row = [Paragraph(str(cell), styles["BodyText"]) for cell in row]
                    table_data.append(wrapped_row)

                # Set minimum column width
                min_width = 50
                available_width = 500
                column_widths = []

                # Calculate column widths based on the longest content in each column
                for col in range(len(data.columns)):
                    max_length = max(len(str(cell)) for cell in data.iloc[:, col])  # Find the max length in the column
                    column_width = max(min_width, max_length * 4)  # 4 is a scaling factor, adjust as needed
                    column_widths.append(column_width)

                # Normalize the total width to fit the available space
                total_width = sum(column_widths)
                if total_width > available_width:
                    scaling_factor = available_width / total_width
                    column_widths = [width * scaling_factor for width in column_widths]

                # Create and style the table
                table = Table(table_data, colWidths=column_widths)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))

                # Create the document with SimpleDocTemplate to handle page breaks
                buffer = BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=letter)

                # List of content for the PDF, including the table and potential page breaks
                story = []
                story.append(table)

                # Check if the content exceeds the page height and needs a page break
                content_height = len(table_data) * 15  # Estimate height based on the number of rows
                page_height = letter[1]  # Page height for letter-sized pages
                if content_height > page_height:
                    story.append(PageBreak())  # Insert a page break if content exceeds page height

                doc.build(story)

                # Return the PDF as a response6 
                buffer.seek(0)
        return HttpResponse(buffer, content_type='application/pdf', headers={'Content-Disposition': f'attachment; filename="{pdf_data.subject}.pdf"'})



# table code that is working and just need to edit 
def generate_quotation(request):
    if request.method == 'POST':
        quotation_form = quotationform(request.POST, request.FILES)
        if quotation_form.is_valid():
            pdf_data = quotation_form.save()
            buffer = BytesIO()
            styles = getSampleStyleSheet()
            styles['BodyText'].fontSize = 12
            styles['BodyText'].leading = 20
            width, height = A4

            # Prepare the content for the document
            story = []

            # Add letterhead
            if pdf_data.letterhead:
                letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)

                def add_header_footer(canvas, doc):
                    canvas.saveState()
                    if pdf_data.letterhead:
                        canvas.drawImage(letterhead_path, 40, height - 100, width=500, height=100)

                    # Add heading
                    canvas.setFont("Helvetica-Bold", 14)
                    canvas.drawString(50, height - 120, "Price Schedule for Goods")
                    
                    # Add additional text fields
                    # canvas.setFont("Helvetica", 12)
                    # canvas.drawString(50, height - 140, f"Date: {pdf_data.created_at.strftime('%d-%m-%Y')}")
                    # canvas.restoreState()

            # Handle Excel table data
            if pdf_data.excel:
                excel_path = os.path.join(settings.MEDIA_ROOT, pdf_data.excel.name)
                data = pd.read_excel(excel_path, engine='openpyxl')
                data = data.where(pd.notnull(data), None)  # Handle NaN values
                data_list = [data.columns.tolist()] + data.values.tolist()

                # Prepare the table data with Paragraphs
                table_data = [
                    [Paragraph(str(cell) if cell else '', styles["BodyText"]) for cell in row]
                    for row in data_list
                ]

                # Calculate column widths
                min_width = 100
                available_width = 500
                column_widths = [max(min_width, len(str(cell)) * 4) for cell in data.columns]

                # Normalize widths
                total_width = sum(column_widths)
                if total_width > available_width:
                    scaling_factor = available_width / total_width
                    column_widths = [width * scaling_factor for width in column_widths]

                # Create the table
                table = Table(table_data, colWidths=column_widths)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                story.append(table)

            # Add a spacer after the table
            story.append(Spacer(1, 20))

            # Add note
            note_paragraph = Paragraph(
                f"<b>Note:</b> {pdf_data.notes}", styles["BodyText"]
            )
            story.append(note_paragraph)

            # Add a spacer before the signature
            story.append(Spacer(1, 50))

            # Add signature
            if pdf_data.signature:
                signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                story.append(Image(signature_path, width=150, height=50))
                story.append(Paragraph("Authorized Signature", styles["BodyText"]))
            else:
                story.append(Paragraph("Authorized Signature: ______________________", styles["BodyText"]))

            # Build the PDF
            doc = SimpleDocTemplate(
                buffer, pagesize=A4,
                rightMargin=30, leftMargin=30,
                topMargin=100, bottomMargin=30
            )
            doc.build(story, onFirstPage=add_header_footer, onLaterPages=add_header_footer)

            # Return the PDF as a response
            buffer.seek(0)
            return HttpResponse(buffer, content_type='application/pdf', headers={'Content-Disposition': f'attachment; filename="finesurgical.pdf" '})
        else:
            quotation_form = quotationform(request.POST, request.FILES)

    else:
        quotation_form = quotationform()

    return render(request, 'pdfgenerator/quotation.html', {'quotation_form': quotation_form})