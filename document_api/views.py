from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from docxtpl import DocxTemplate
from datetime import date
import datetime
import os
import win32com.client
import pythoncom
from django.views.decorators.http import require_POST


@require_POST
@csrf_exempt
def generate_document(request):
    data = request.POST.dict()

    sex = data.get('sex', 'Mr.')
    dropdd = data.get('dropdd', 'TOO')
    entry = data.get('entry', '')
    entry2 = data.get('entry2', '')
    entry3 = data.get('entry3', '')
    casenumber = data.get('casenumber', '')
    exchanger = data.get('exchanger', '')
    gasfee = data.get('gasfee', '')
    ethaddress = data.get('ethaddress', '')
    gasdeposit = data.get('gasdeposit', '')
    callername = data.get('callername', '')

    if dropdd == 'TOO':
        doc = DocxTemplate("TOO.docx")
    else:
        doc = DocxTemplate("TD.docx")

    old_word = '$name'
    dropdown_result = data.get('dropdown', '')
    Name = entry
    Surname = entry2
    amount = entry3
    case_number = casenumber
    exchanger_new = exchanger
    gasfee_new = gasfee
    eth_address_new = ethaddress
    gas_deposit_new = gasdeposit
    Callername = callername
    old_word5 = 'Sex'
    old_word1 = '$surname'
    old_casenumber = '$casenumber'
    exchangerr = '$exchanger'
    gasfeetext = 'gasfeetext'
    eth_address = 'ethaddresstext'
    gas_deposit = 'gasdeposittext'
    old_word2 = 'datetimenowand3days'
    Date = date.today()
    Date3days = datetime.datetime.today() + datetime.timedelta(days=3)
    old_word3 = 'amounttoreceive'

    if sex.lower() == "man":
        sex = "Mr"
    elif sex.lower() == "woman":
        sex = "Mrs"

    if dropdd == 'TOO':
        context = {
            'casenumber': case_number,
            'datetime': Date.strftime("%d %b") + '-' + Date3days.strftime("%d %b"),
            'amounttoreceive': '$' + amount,
            'gasfeetext': '$' + gasfee_new,
            'exchanger': exchanger_new,
            'name': Name,
            'surname': Surname,
            'gasdeposit': '$' + gas_deposit_new,
            'ethaddress': eth_address_new
        }
    else:
        context = {
            'casenumber': case_number,
            'datetime': Date.strftime("%d %b"),
            'sex': sex,
            'amounttoreceive': '$' + amount,
            'name': Name,
            'surname': Surname,
            'callername': Callername,
        }

    doc.render(context)

    output_filename = f'{Name}_{Surname}.docx'
    output_path = os.path.join('media', output_filename)
    doc.save(output_path)

    wdFormatPDF = 17
    in_file = os.path.abspath(output_path)
    out_file = os.path.abspath(f'{Name}_{Surname}.pdf')

    word = win32com.client.Dispatch("Word.Application",pythoncom.CoInitialize())


    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

    return JsonResponse({'message': 'Document generated successfully.', 'url': f'/{Name}_{Surname}.pdf'})

