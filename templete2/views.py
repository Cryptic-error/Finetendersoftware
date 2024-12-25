
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
from .forms import PDFForm,quotationform,pricescheduleform
from django.shortcuts import render
from reportlab.lib.enums import TA_JUSTIFY,TA_CENTER,TA_RIGHT
from reportlab.platypus import Spacer
from reportlab.platypus import Image



from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth

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
    numberofrow = 0
    
    if request.method == 'POST':
        pdfdata_form = PDFForm(request.POST, request.FILES)
        if pdfdata_form.is_valid():
            pdf_data = pdfdata_form.save()
            buffer = BytesIO()
            p = canvas.Canvas(buffer, pagesize=A4)
            width, height = A4
            styles = getSampleStyleSheet()
            styles['BodyText'].fontSize = 12
            styles['BodyText'].leading = 20
            styles['BodyText'].alignment = TA_JUSTIFY
# style normal
            styles['Normal'].leading = 20
            styles['Normal'].fontSize = 12
            styles['Normal'].alignment = TA_RIGHT
            # words = rupee_format(pdf_data.amount)

            if pdfdata_form.cleaned_data.get('include_selfdeclarationletter'):

                # Draw letterhead image at the top, if available
                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 10, height - 100, width=600, height=100)


                # Add the address and content text
                elements = []
                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    datefields = f"""
                        Date: {date}<br/>
                        Id: {pdf_data.id_no}<br/>
                        Name: {pdf_data.nameofcontract}
                    """
                    styles['Normal'].alignment = TA_RIGHT
                    elements.append(Paragraph(datefields, styles['Normal']))
                    styles['BodyText'].alignment = TA_JUSTIFY
                # Address block as a Paragraph
                if pdf_data.designation:
                    address = f"""
                        To,<br/>
                        {pdf_data.designation or ""}<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))
                else:
                    address = f"""
                        To,<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))
                # Declaration Title
                elements.append(Paragraph("<u><b>Self Declaration Letter</b></u>", styles['Title']))

                # Body Content
                body_text = """
                    Dear Sir/Madam,<br/>
                    &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;This refers to you that related to above subject with respect to the procurement procedure of 
                    the organization. We Sharvil Energy Pvt. Ltd. declare that we are eligible to participate 
                    in the bidding process.<br/> We have no conflict of interest with Bidder with respect to the proposed 
                    bid procurement proceedings and have no pending Litigation with the bidder.<br/><br/>
                    Thank you, <br/> Sincerely yours,
                """
                elements.append(Paragraph(body_text, styles['BodyText']))

                y_position = height - 120  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position, width=300, height=80)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60
    
                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.ourdesignation)
                if pdf_data.footer:
                    y_position=height-771
                    print(y_position)
                    print(height)
                    footer_path = os.path.join(settings.MEDIA_ROOT, pdf_data.footer.name)
                    p.drawImage(footer_path, 0, y_position - 60, width=600, height=80)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60
                p.showPage()  # Move to the next page if needed

            if pdfdata_form.cleaned_data.get('include_warrantyletter'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 10, height - 100, width=600, height=100)

                elements = []
                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    datefields = f"""
                        Date: {date}<br/>
                        Id: {pdf_data.id_no}<br/>
                        Name: {pdf_data.nameofcontract}
                    """
                    styles['Normal'].alignment = TA_RIGHT
                    elements.append(Paragraph(datefields, styles['Normal']))
                    styles['BodyText'].alignment = TA_JUSTIFY

            # Warranty commitment letter page
                # Address block as a Paragraph
                if pdf_data.designation:
                    address = f"""
                        To,<br/>
                        {pdf_data.designation or ""}<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))
                else:
                    address = f"""
                        To,<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))

               

                # Creating the paragraphs for warranty commitment
                

                elements.append(Paragraph("<u><b>Warranty/Commitment Letter</b></u>", styles['Title']))

                commitment_text = f"""
                    Dear Sir/Madam,<br/>
                    &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;This refers to your Tender for <b>The Procurement of {pdf_data.subject}.</b> We, Sharvil Energy Pvt. Ltd., declare that we are eligible to participate 
                    in the bidding process.if awarded for the tender we warrant that the Goods are new, unused &amp; of
                    the most recent or current models &amp; that they incorporate all recent improvements in design &amp; materials
                    unless provided otherwise in the Contract.<br/>
                    We further warrant that the Goods shall be free from defects arising from any act or commission of us or
                    arising from design, materials &amp; workmanship under normal use in the conditions prevailing in the country
                    of final destination.<br/>
                    The comprehensive warranty &amp; service warranty for offered product shall remain valid for the required
                    periods after the Goods or any portion thereof as the case may be, have been delivered to &amp; accepted at the
                    final destination.<br/>
                    We further ensured that during the warranty period we will provide Preventive Maintenance (PPM) along
                    with corrective / breakdown maintenance whenever required.<br/><br/>
                    Thank you,<br/> Sincerely yours,<br/>
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 110  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 10  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 60, width=250, height=70)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.ourdesignation)
                if pdf_data.footer:
                    y_position=height-771
                    print(y_position)
                    print(height)
                    footer_path = os.path.join(settings.MEDIA_ROOT, pdf_data.footer.name)
                    p.drawImage(footer_path, 0, y_position - 60, width=600, height=80)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60
                p.showPage()
            

            if pdfdata_form.cleaned_data.get('include_manufactureletter'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 10, height - 100, width=600, height=100)

                elements = []
                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    datefields = f"""
                        Date: {date}<br/>
                        Id: {pdf_data.id_no}<br/>
                        Name: {pdf_data.nameofcontract}
                    """
                    styles['Normal'].alignment = TA_RIGHT
                    elements.append(Paragraph(datefields, styles['Normal']))
                    styles['BodyText'].alignment = TA_JUSTIFY

            # Warranty commitment letter page
                # Address block as a Paragraph
                if pdf_data.designation:
                    address = f"""
                        To,<br/>
                        {pdf_data.designation or ""}<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))
                else:
                    address = f"""
                        To,<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))

                

                # Creating the paragraphs for warranty commitment
                

                elements.append(Paragraph("<u><b>Date of Manufacture and Brand New Machine</b></u>", styles['Title']))

                commitment_text = f"""
                    Dear Sir/Madam,<br/>
                    &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;This refers to your tender for <b> The Procurement of {pdf_data.subject}</b>. We, Sharvil Energy Pvt. Ltd., declare that the offered equipment shall
                    be brand new manufactured after placement of the order. <br/><br/>             
                    We shall also submit certification with the date of manufacture of the machine along with shipment
                    of the system..<br/><br/>
                    Thank you, <br/>Sincerely yours,<br/>
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 120  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=80)
                    y_position -= 100  # Adjust y position below the signature image
                else:
                    y_position -=100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.ourdesignation)
                if pdf_data.footer:
                    y_position=height-771
                    print(y_position)
                    print(height)
                    footer_path = os.path.join(settings.MEDIA_ROOT, pdf_data.footer.name)
                    p.drawImage(footer_path, 0, y_position - 60, width=600, height=80)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60
                p.showPage()


            if pdfdata_form.cleaned_data.get('include_deliverycommitmentletter'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 10, height - 100, width=600, height=100)

                elements = []
                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    datefields = f"""
                        Date: {date}<br/>
                        Id: {pdf_data.id_no}<br/>
                        Name: {pdf_data.nameofcontract}
                    """
                    styles['Normal'].alignment = TA_RIGHT
                    elements.append(Paragraph(datefields, styles['Normal']))
                    styles['BodyText'].alignment = TA_JUSTIFY

            # Warranty commitment letter page
                # Address block as a Paragraph
                if pdf_data.designation:
                    address = f"""
                        To,<br/>
                        {pdf_data.designation or ""}<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))
                else:
                    address = f"""
                        To,<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))

            

                # Creating the paragraphs for warranty commitment
                

                elements.append(Paragraph("<u><b>Delivery commitment letter</b></u>", styles['Title']))

                commitment_text = f"""
                    Dear Sir/Madam,<br/>
                    &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;This refers to your tender for <b> The Procurement of {pdf_data.subject}</b>. We, Sharvil Energy Pvt. Ltd., declare that We Fine Surgicals Nepal Pvt.Ltd., if we are awarded the
                    contract, the Machine can be delivered within <b>{pdf_data.days} days</b> after conformation.<br/><br/>
                    Thank you, <br/> Sincerely yours,<br/>
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 120  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=80)
                    y_position -= 100  # Adjust y position below the signature image
                else:
                    y_position -=100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.ourdesignation)
                if pdf_data.footer:
                    y_position=height-771
                    print(y_position)
                    print(height)
                    footer_path = os.path.join(settings.MEDIA_ROOT, pdf_data.footer.name)
                    p.drawImage(footer_path, 0, y_position - 60, width=600, height=80)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60
                p.showPage()


            # installation 

            if pdfdata_form.cleaned_data.get('include_installationletter'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 10, height - 100, width=600, height=100)

                elements = []
                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    datefields = f"""
                        Date: {date}<br/>
                        Id: {pdf_data.id_no}<br/>
                        Name: {pdf_data.nameofcontract}
                    """
                    styles['Normal'].alignment = TA_RIGHT
                    elements.append(Paragraph(datefields, styles['Normal']))
                    styles['BodyText'].alignment = TA_JUSTIFY

            # Warranty commitment letter page
                # Address block as a Paragraph
                if pdf_data.designation:
                    address = f"""
                        To,<br/>
                        {pdf_data.designation or ""}<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))
                else:
                    address = f"""
                        To,<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))

                

                # Creating the paragraphs for warranty commitment
                

                elements.append(Paragraph("<u><b>Installation Demonstration And Training </b></u>", styles['Title']))

                commitment_text = f"""
                    Dear Sir/Madam,<br/>
                    &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;This refers to your tender for <b> The Procurement of {pdf_data.subject}</b>. We, Sharvil Energy Pvt. Ltd. If we are awarded the contract installation &amp;
                     demonstration of the equipment shall be carried out by the trained engineers.<br/>They shall also impart training to the engineer, technicians &amp; users at site during Installation of the system.<br/>
                Thank you, <br/> Sincerely yours<br/>
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 120  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=80)
                    y_position -= 100  # Adjust y position below the signature image
                else:
                    y_position -=100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.ourdesignation)
                if pdf_data.footer:
                    y_position=height-771
                    print(y_position)
                    print(height)
                    footer_path = os.path.join(settings.MEDIA_ROOT, pdf_data.footer.name)
                    p.drawImage(footer_path, 0, y_position - 60, width=600, height=80)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60
                p.showPage()

            if pdfdata_form.cleaned_data.get('include_attorneyletter'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 10, height - 100, width=600, height=100)

                elements = []
                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    datefields = f"""
                        Date: {date}<br/>
                        Id: {pdf_data.id_no}<br/>
                        Name: {pdf_data.nameofcontract}
                    """
                    styles['Normal'].alignment = TA_RIGHT
                    elements.append(Paragraph(datefields, styles['Normal']))
                    styles['BodyText'].alignment = TA_JUSTIFY

            # Warranty commitment letter page
                # Address block as a Paragraph
                if pdf_data.designation:
                    address = f"""
                        To,<br/>
                        {pdf_data.designation or ""}<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))
                else:
                    address = f"""
                        To,<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))

                

                # Creating the paragraphs for warranty commitment
                

                elements.append(Paragraph("<u><b>Power of attorney </b></u>", styles['Title']))

                commitment_text = f"""
                    Dear Sir/madam,<br/>
                    &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;We, Sharvil Energy Pvt. Ltd. Here by to assign {pdf_data.proprietor_name} {pdf_data.ourdesignation} of this company to handle all tender and quotation related to <b> The Procurement of {pdf_data.subject}</b>.<br/> He is authorized to do all necessary documentation on our behalf and purpose.<br/> 
                  Thank you, <br/> Sincerely yours
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 120  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.attornitysign:
                    attornitysign_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(attornitysign_path, 50, y_position - 80, width=300, height=80)
                    y_position -= 100  # Adjust y position below the signature image
                else:
                    y_position -=100  # Adjust y position below the signature image
 # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, f"{pdf_data.attornity_name}")
                p.drawString(50, y_position - 40, f"{pdf_data.attornity_designation}")
                if pdf_data.footer:
                    y_position=height-771
                    print(y_position)
                    print(height)
                    footer_path = os.path.join(settings.MEDIA_ROOT, pdf_data.footer.name)
                    p.drawImage(footer_path, 0, y_position - 60, width=600, height=80)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60
                p.showPage()

            if pdfdata_form.cleaned_data.get('include_bidsubmissionform'):
                
                words = rupee_format(int(pdf_data.amount))
                print("Formatted Words:", words)

                y_position = height - 160 
                
                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 10, height - 100, width=600, height=100)
                elements = []
                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    datefields = f"""
                        Date: {date}<br/>
                        Id: {pdf_data.id_no}<br/>
                        Name: {pdf_data.nameofcontract}
                    """
                    styles['Normal'].alignment = TA_RIGHT
                    elements.append(Paragraph(datefields, styles['Normal']))
                    styles['BodyText'].alignment = TA_JUSTIFY
                    

            # Warranty commitment letter page
                # Address block as a Paragraph
                if pdf_data.designation:
                    address = f"""
                        To,<br/>
                        {pdf_data.designation or ""}<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))
                else:
                    address = f"""
                        To,<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))

                

                # Creating the paragraphs for warranty commitment
                

                elements.append(Paragraph("<u><b>Bid submission Forms </b></u>", styles['Title']))

                commitment_text = f"""
                    we,the undersigned declare that,<br/>
                    (a) We have examined &amp; have no reservations to the Bidding Documents, including Addenda No.: No addenda.<br/><br/>
                    (b) We offer to supply in conformity with the Bidding Documents &amp; in accordance with the Delivery Schedules
                    specified in the Schedule of Requirements the following Goods &amp; related Services:{pdf_data.subject} . The total
                    price of our Bid, excluding any discounts offered in item (d) below, is; <b> {words} (Nrs)</b>
                    excluding Tax and Vat.<br/><br/>
                    (c) The discounts offered and the methodology for their application are: Not applicable.<br/><br/>
                    (d) Our bid shall be valid for the period of {pdf_data.bidvaliditydays} days from the date fixed for the bid submission deadline in
                    accordance with the Bidding Document, &amp; it shall remain binding upon us &amp; may be accepted at any time
                    before the expiration of that period<br/><br/>
                    (e) If our Bid is accepted, we commit to obtain a Performance Security in the amount as specified in 1TB 41 for
                    the due performance of the Contract<br/><br/>
                    (f) We are not participating, as Bidders, in more than one Bid in this bidding process, other than alternative
                    offers in accordance with the Biding Document.<br/><br/>
                    
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))
                
                
                y_position = height - 120  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20
                if pdf_data.footer:
                    y_position=height-771
                    print(y_position)
                    print(height)
                    footer_path = os.path.join(settings.MEDIA_ROOT, pdf_data.footer.name)
                    p.drawImage(footer_path, 0, y_position - 60, width=600, height=100)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60
                
                p.showPage()

            if pdfdata_form.cleaned_data.get('include_bidsubmissionform'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 10, height - 100, width=600, height=100)



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

                y_position = height - 120  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 10  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=80)
                    y_position -= 100  # Adjust y position below the signature image
                else:
                    y_position -=100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.ourdesignation)
                if pdf_data.footer:
                    y_position=height-771
                    print(y_position)
                    print(height)
                    footer_path = os.path.join(settings.MEDIA_ROOT, pdf_data.footer.name)
                    p.drawImage(footer_path, 0, y_position - 60, width=600, height=80)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60
                p.showPage()

# letter of technicxal bid

            if pdfdata_form.cleaned_data.get('include_technicalbid'):

                y_position = height - 200 
                
                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 10, height - 100, width=600, height=100)

                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    datefields = f"""
                        Date: {date}<br/>
                        Id: {pdf_data.id_no}<br/>
                        Name: {pdf_data.nameofcontract}
                    """
                    styles['Normal'].alignment = TA_RIGHT
                    elements.append(Paragraph(datefields, styles['Normal']))
                    styles['BodyText'].alignment = TA_JUSTIFY

            # Warranty commitment letter page
                # Address block as a Paragraph
                if pdf_data.designation:
                    address = f"""
                        To,<br/>
                        {pdf_data.designation or ""}<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))
                else:
                    address = f"""
                        To,<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))

                elements = []

                # Creating the paragraphs for warranty commitment
                

                elements.append(Paragraph("<u><b>Letter of Technical Bid </b></u>", styles['Title']))

                commitment_text = f"""
                    (a) We have examined and have no reservations to the Bidding Document, including Addenda issued in
                    accordance with Instructions to Bidders (ITB) Clause 9;<br/><br/>
                    (b) We offer to supply in conformity with the Bidding Document and in accordance with the delivery
                    schedule specified in the Section V (Schedule of Requirements), the following Goods and Related
                    Services: <b>The Procurement of {pdf_data.subject}.</b><br/><br/>
                    (c) Our Bid consisting of the Technical Bid and the Price Bid shall be valid for a period of 90 days from the
                    date fixed for the bid submission deadline in accordance with the Bidding Document, and it shall
                    remain binding upon us and may be accepted at any time before the expiration of that period;<br/><br/>
                    (d) Our firm, including any subcontractors or suppliers for any part of the Contract, has nationalities from
                    eligible countries in accordance with ITB 4.8 and meets the requirements of ITB 3.4 &amp; 3.5;<br/><br/>
                    (e) We are not participating, as a Bidder or as a subcontractor/supplier, in more than one Bid in this bidding
                    process in accordance with ITB 4.3(e), other than alternative Bids in accordance with ITB 14;<br/><br/>
                    (f) Our firm, its affiliates or subsidiaries, including any Subcontractors or Suppliers for any part of the
                    contract, has not been declared ineligible by DP, under the Purchaser’s country laws or official
                    regulations or by an act of compliance with a decision of the United Nations Security Council;<br/>
                    
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))
                
                
                y_position = height - 120  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 10
                
                if pdf_data.footer:
                    y_position=height-771
                    print(y_position)
                    print(height)
                    footer_path = os.path.join(settings.MEDIA_ROOT, pdf_data.footer.name)
                    p.drawImage(footer_path, 0, y_position - 60, width=600, height=80)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60 
                p.showPage()

            if pdfdata_form.cleaned_data.get('include_technicalbid'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 10, height - 100, width=600, height=100)



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

                    Amount<br/><br/>

                    &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; &#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; NONE<br/><br/></b>

                    (j) We declare that we are solely responsible for the authenticity of the documents submitted by us. The
                        document and information submitted by us are true and correct. If any document/information given is
                        found to be concealed at a later date, we shall accept any legal actions by the purchaser.<br/><br/>
                    (k) We agree to permit GoN/DP or its representative to inspect our accounts and records and other
                        documents relating to the bid submission and to have them audited by auditors appointed by the
                        GoN/DP.<br/>
                        Thank you,<br/>Sincerely yours
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 100  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=80)
                    y_position -= 100  # Adjust y position below the signature image
                else:
                    y_position -=100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, pdf_data.proprietor_name)
                p.drawString(50, y_position - 40, pdf_data.ourdesignation)
                if pdf_data.footer:
                    y_position=height-771
                    print(y_position)
                    print(height)
                    footer_path = os.path.join(settings.MEDIA_ROOT, pdf_data.footer.name)
                    p.drawImage(footer_path, 0, y_position - 60, width=600, height=80)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60
                p.showPage()

            if pdfdata_form.cleaned_data.get('include_pricebid'):
                vats_amt = round(pdf_data.amount * 1.13, 2)
                vat_words = rupee_format(int(vats_amt))
                words = rupee_format(int(pdf_data.amount))
                print("Formatted Words:", words)
                y_position = height - 200 
                
                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 10, height - 100, width=600, height=100)
                elements = []
                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    datefields = f"""
                        Date: {date}<br/>
                        Id: {pdf_data.id_no}<br/>
                        Name: {pdf_data.nameofcontract}
                    """
                    styles['Normal'].alignment = TA_RIGHT
                    elements.append(Paragraph(datefields, styles['Normal']))
                    styles['BodyText'].alignment = TA_JUSTIFY

            # Warranty commitment letter page
                # Address block as a Paragraph
                if pdf_data.designation:
                    address = f"""
                        To,<br/>
                        {pdf_data.designation or ""}<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))
                else:
                    address = f"""
                        To,<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))

                

                # Creating the paragraphs for warranty commitment
                

                elements.append(Paragraph("<u><b>Letter of Price Bid </b></u>", styles['Title']))

                commitment_text = f"""
                    (a) We have examined and have no reservations to the Bidding Document, including Addenda issued in
                    accordance with Instructions to Bidders (ITB) Clause 9;<br/><br/>
                    (b) We offer to supply in conformity with the Bidding Document and in accordance with the delivery
                    schedule specified in the Section V (Schedule of Requirements), the following Goods and Related
                    Services: <b>The Procurement of {pdf_data.subject}.</b><br/><br/>
                    (c) The total price of our Bid, excluding any discounts offered in item (d) below, is: {pdf_data.amount}(In Words {vat_words})<br/><br/>
                    (d) The discounts offered and the methodology for their application are:<br/>
                    - The discounts offered are  NONE<br/>
                    - The exact method of calculations to determine the net price after application of discounts is shown below
                     .NONE<br/><br/>
                    (e) Our bid shall be valid for a period of {pdf_data.bidvaliditydays} days from the date fixed for the bid submission deadline in accordance with the Bidding Documents, and it shall remain binding upon us and may be accepted at any time before the expiration of that period.<br/><br/>
                    (f) If our bid is accepted, we commit to obtain a performance security in accordance with the Bidding Document.<br/><br/>
                    
                                        
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))
                
                
                y_position = height - 110  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 10
                
                if pdf_data.footer:
                    y_position=height-771
                    print(y_position)
                    print(height)
                    footer_path = os.path.join(settings.MEDIA_ROOT, pdf_data.footer.name)
                    p.drawImage(footer_path, 0, y_position - 60, width=600, height=80)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60
                
                p.showPage()

            if pdfdata_form.cleaned_data.get('include_pricebid'):

                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 10, height - 100, width=600, height=100)



                elements = []


                commitment_text = f"""
                    (g) We understand that this bid, together with your written acceptance thereof included in your notification of award, shall constitute a binding contract between us, until a formal contract is prepared and executed.<br/><br/>
                    (h) We understand that you are not bound to accept the lowest evaluated bid or any other bid that you may receive. <br/><br/>
                    (i) We agree to permit the Employer/DP 1 or its representative to inspect our accounts and records and other documents relating to the bid submission and to have them audited by auditors 2 appointed by the Employer. <br/><br/>
                    (j) We confirm and stand by our commitments and other declarations made in connection with the submission of our Letter of Technical Bid.<br/><br/>
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))

                y_position = height - 120  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20  # Adjust for the next element

                p.drawString(50, y_position - 20, f"Name:{pdf_data.proprietor_name}")
                p.drawString(50, y_position - 40, f"In the capacity of :{pdf_data.ourdesignation}")
                p.drawString(50, y_position - 60, "signed")
                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=80)
                    y_position -= 100  # Adjust y position below the signature image
                else:
                    y_position -=100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, "Duly authorized to sign the bid for & on behalf of: Fine Surgicals Nepal Pvt. Ltd.")
                p.drawString(50, y_position - 40, f"Date:{pdf_data.date}")
                if pdf_data.footer:
                    y_position=height-771
                    print(y_position)
                    print(height)
                    footer_path = os.path.join(settings.MEDIA_ROOT, pdf_data.footer.name)
                    p.drawImage(footer_path, 0, y_position - 60, width=600, height=80)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60
                p.showPage()



            if pdfdata_form.cleaned_data.get('include_quotationandprice'):
                
                vats_amt = round(pdf_data.amount * 1.13, 2)
                vat_words = rupee_format(int(vats_amt))
                words = rupee_format(int(pdf_data.amount))
                print("Formatted Words:", words)
                
                y_position = height - 200 
                
                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    p.drawImage(letterhead_path, 10, height - 100, width=600, height=100)

                if pdf_data.date:
                    date = pdf_data.date.strftime("%Y-%m-%d")
                    datefields = f"""
                        Date: {date}<br/>
                        Id: {pdf_data.id_no}<br/>
                        Name: {pdf_data.nameofcontract}
                    """
                    styles['Normal'].alignment = TA_RIGHT
                    elements.append(Paragraph(datefields, styles['Normal']))
                    styles['BodyText'].alignment = TA_JUSTIFY
                   
            # Warranty commitment letter page
                # Address block as a Paragraph
                if pdf_data.designation:
                    address = f"""
                        To,<br/>
                        {pdf_data.designation or ""}<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))
                else:
                    address = f"""
                        To,<br/>
                        {pdf_data.institution_name or ""}<br/>
                        {pdf_data.address or ""}
                    """
                    elements.append(Paragraph(address, styles['BodyText']))

                elements = []

                # Creating the paragraphs for warranty commitment
                

                elements.append(Paragraph("<u><b>Quotation and Price Schedule</b></u>", styles['Title']))

                commitment_text = f"""
                    Having examined the Sealed Quotation (SQ) documents, we the undersigned, offer supply and delivery of
                    {pdf_data.subject} in conformity with the said SQ documents for the sum of or
                    such other sums <u><b>{vats_amt}.</b></u> (In Words <u><b> {vat_words})</b></u>as may be ascertained in accordance with the Schedule of
                    Prices attached herewith and made part of this SQ.
                    We undertake, if our SQ is accepted, to deliver the goods in accordance with the delivery schedule specified
                    in the Schedule of Requirements.<br/><br/>
                    If our SQ is accepted, we will obtain the guarantee of bank if mentioned in contract due performance of the Contract, in the form prescribed by the Purchaser.<br/>
                    We agree to abide by this SQ for a Period of {pdf_data.bidvaliditydays} days from the date fixed for SQ opening it shall remain
                    binding upon us and may be accepted at any time before the expiration of that period.<br/>
                    Until a formal Contract is prepared and executed, this SQ, together with your written acceptance thereof
                    and your notification of award, shall constitute a binding Contract between us.<br/>
                    We understand that you are not bound to accept the lowest or any SQ you may receive.<br/>
                    Dated this <u> {pdf_data.datedthis}</u><br/><br/><br/>
                    
                    
                """
                elements.append(Paragraph(commitment_text, styles['BodyText']))
                
                
                y_position = height - 120  # Start position
                for element in elements:
                    w, h = element.wrap(width - 100, height)
                    element.drawOn(p, 50, y_position - h)
                    y_position -= h + 20
                p.drawString(50, y_position - 20, f"Name:{pdf_data.proprietor_name}")
                p.drawString(50, y_position - 40, f"In the capacity of :{pdf_data.ourdesignation}")
                p.drawString(50, y_position - 60, "signed")
                # Add signature at a specific position if available
                if pdf_data.signature:
                    signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                    p.drawImage(signature_path, 50, y_position - 80, width=300, height=80)
                    y_position -= 100  # Adjust y position below the signature image
                else:
                    y_position -=100  # Adjust y position below the signature image

                line_y_position = y_position - 5  # Adjust this value for spacing below the image
                p.setLineWidth(1)  # Set the line width if needed
                p.line(40, line_y_position, width-300, line_y_position)

                p.drawString(50, y_position - 20, "Duly authorized to sign the bid for & on behalf of: Sharvil Energy Pvt. Ltd.")
                p.drawString(50, y_position - 40, f"Date:{pdf_data.date}")
                if pdf_data.footer:
                    y_position=height-771
                    print(y_position)
                    print(height)
                    footer_path = os.path.join(settings.MEDIA_ROOT, pdf_data.footer.name)
                    p.drawImage(footer_path, 0, y_position - 60, width=600, height=80)
                    y_position -= 60  # Adjust y position below the signature image
                else:
                    y_position -= 60
                
                p.showPage()
    


            p.save()

            # Return the PDF as a response
            buffer.seek(0)
            return HttpResponse(buffer, content_type='application/pdf', headers={'Content-Disposition': f'attachment; filename="{pdf_data.subject}.pdf" ', 'rowcount':str(numberofrow)})

    else:
        pdfdata_form = PDFForm()

    return render(request, 'pdfgenerator/forms.html', {'pdfdata_form': pdfdata_form})



def calculate_column_widths(data, font_name="Helvetica", font_size=10, min_width=50, max_width=200):
    column_widths = []
    for col in data.columns:
        max_cell_width = max(
            stringWidth(str(cell), font_name, font_size) if cell else 0
            for cell in data[col]
        )
        column_widths.append(min(max(max_cell_width + 10, min_width), max_width))
    return column_widths





def generate_priceschedule(request):
    if request.method == 'POST':
        pricescheduleform_form = pricescheduleform(request.POST, request.FILES)
        if pricescheduleform_form.is_valid():
            pdf_data = pricescheduleform_form.save()
            buffer = BytesIO()
            styles = getSampleStyleSheet()
            centered_heading_style = styles['Heading3']
            centered_heading_style.alignment = TA_CENTER
            styles['BodyText'].fontSize = 10
            styles['BodyText'].leading = 20
            width, height = A4

            # Prepare the content for the document
            story = []

            # Function to add header/footer
            def add_header_footer(canvas, doc):
                canvas.saveState()
                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    canvas.drawImage(letterhead_path, 40, height - 100, width=500, height=80)
                canvas.restoreState()

            # Add "Price Schedule for Goods" heading
            heading = Paragraph(f"<u>Price Schedule for Goods</u> <br/><br/> Name of bidder : {pdf_data.institutionname}", centered_heading_style)
            story.append(heading)
            story.append(Spacer(1, 20))

            # Handle Excel table data
            if pdf_data.excel:
                excel_path = os.path.join(settings.MEDIA_ROOT, pdf_data.excel.name)
                data = pd.read_excel(excel_path, engine='openpyxl')
                data = data.where(pd.notnull(data), None)  # Handle NaN values
                data_list = [data.columns.tolist()] + data.values.tolist()
                vats_amt = (pdf_data.amount*1.13)
                vat_words = rupee_format(int(vats_amt))
                words = rupee_format(int(pdf_data.amount))
                print("Formatted Words:", words)
                # Prepare the table data with Paragraphs
                table_data = [
                    [Paragraph(str(cell) if cell else '', styles["BodyText"]) for cell in row]
                    for row in data_list
                ]

                # Calculate column widths
                # Calculate column widths dynamically
                column_widths = calculate_column_widths(data, font_name="Helvetica", font_size=10)

                # Normalize widths to fit within the page
                available_width = width - 60  # Account for page margins
                total_width = sum(column_widths)
                if total_width > available_width:
                    scaling_factor = available_width / total_width
                    column_widths = [w * scaling_factor for w in column_widths]

                # Prepare table data
                table_data = [
                    [Paragraph(str(cell) if cell else '', styles["BodyText"]) for cell in row]
                    for row in data_list
                ]

                # Create table
                table = Table(table_data, colWidths=column_widths)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                story.append(table)


            # Add a spacer after the table
            story.append(Spacer(1, 10))
            story.append(Paragraph(f"In Words :{vat_words}"))
            story.append(Paragraph(f"Notes*:{pdf_data.notes}"))

            # Add signature section
            # story.append(Spacer(1, 50))
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
            return HttpResponse(buffer, content_type='application/pdf', headers={'Content-Disposition': f'attachment; filename="finesurgical.pdf"'})

    else:
        pricescheduleform_form = pricescheduleform()

    return render(request, 'pdfgenerator/priceschedule.html', {'pricescheduleform_form': pricescheduleform_form})

def generate_quotation(request):
    if request.method == 'POST':
        quotation_form = quotationform(request.POST, request.FILES)
        if quotation_form.is_valid():
            pdf_data = quotation_form.save()
            # p = canvas.Canvas(buffer, pagesize=A4)
            buffer = BytesIO()
            styles = getSampleStyleSheet()
            centered_heading_style = styles['Heading3']
            centered_heading_style.alignment = TA_CENTER
            styles['BodyText'].fontSize = 10
            styles['BodyText'].leading = 20
            width, height = A4

            # Prepare the content for the document
            story = []

            # Function to add header/footer
            def add_header_footer(canvas, doc):
                canvas.saveState()
                if pdf_data.letterhead:
                    letterhead_path = os.path.join(settings.MEDIA_ROOT, pdf_data.letterhead.name)
                    canvas.drawImage(letterhead_path, 40, height - 100, width=500, height=80)
                canvas.restoreState()

            address = f"""
                To,<br/>
                {pdf_data.designation or ""}<br/>
                {pdf_data.institution_name or ""}<br/>
                {pdf_data.address or ""}<br/>
                
        """
            story.append(Paragraph(address, styles['BodyText']))

            story.append(Paragraph(f"<u><b>subject </b></u>:{pdf_data.subject}", styles['Heading3']))
            story.append(Paragraph(f"Dear sir/madam<br/>&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160; {pdf_data.paragraph}", styles['BodyText']))


            # Handle Excel table data
            if pdf_data.excel:
                excel_path = os.path.join(settings.MEDIA_ROOT, pdf_data.excel.name)
                data = pd.read_excel(excel_path, engine='openpyxl')
                data = data.where(pd.notnull(data), None)  # Handle NaN values
                data_list = [data.columns.tolist()] + data.values.tolist()
                vats_amt = (pdf_data.amount*1.13)
                vat_words = rupee_format(int(vats_amt))
                words = rupee_format(int(pdf_data.amount))
                print("Formatted Words:", words)
                # Prepare the table data with Paragraphs
                table_data = [
                    [Paragraph(str(cell) if cell else '', styles["BodyText"]) for cell in row]
                    for row in data_list
                ]

                # Calculate column widths
                # Calculate column widths dynamically
                column_widths = calculate_column_widths(data, font_name="Helvetica", font_size=10)

                # Normalize widths to fit within the page
                available_width = width - 60  # Account for page margins
                total_width = sum(column_widths)
                if total_width > available_width:
                    scaling_factor = available_width / total_width
                    column_widths = [w * scaling_factor for w in column_widths]

                # Prepare table data
                table_data = [
                    [Paragraph(str(cell) if cell else '', styles["BodyText"]) for cell in row]
                    for row in data_list
                ]

                # Create table
                table = Table(table_data, colWidths=column_widths)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                story.append(table)


            # Add a spacer after the table
            story.append(Spacer(1, 10))
            story.append(Paragraph(f"In Words :{vat_words}"))
            story.append(Paragraph(f"Notes*:{pdf_data.notes}"))

            # Add signature section
            story.append(Spacer(1, 20))
            if pdf_data.signature:
                signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
                story.append(Image(signature_path, width=100, height=50, hAlign='LEFT'))

                story.append(Paragraph("Authorized Signature", styles["BodyText"]))
            else:
                story.append(Paragraph("Authorized Signature: ______________________", styles["BodyText"]))
            # if pdf_data.signature:
            #     canvas.saveState()
            #     signature_path = os.path.join(settings.MEDIA_ROOT, pdf_data.signature.name)
            #     # canvas.drawImage(signature_path, 50, y_position - 80, width=300, height=80)
            #     canvas.drawImage(signature_path, 50, height - 180, width=300, height=80)

            #     y_position -= 100  # Adjust y position below the signature image
            # else:
            #     y_position -=100 
            # Build the PDF
            doc = SimpleDocTemplate(
                buffer, pagesize=A4,
                rightMargin=30, leftMargin=30,
                topMargin=100, bottomMargin=30
            )
            doc.build(story, onFirstPage=add_header_footer, onLaterPages=add_header_footer)

            # Return the PDF as a response
            buffer.seek(0)
            return HttpResponse(buffer, content_type='application/pdf', headers={'Content-Disposition': f'attachment; filename="finesurgical.pdf"'})

    else:
        quotation_form = quotationform()

    return render(request, 'pdfgenerator/quotation.html', {'quotation_form': quotation_form})


