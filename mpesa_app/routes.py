from mpesa_app import app
from flask import render_template, url_for, request, make_response
from mpesa_app.utils import extract_from_pdf, parse_mpesa_content, find_name, paidin, withdrawal, listing, dfs_tabs

@app.route('/', methods=['GET', 'POST'])
def index():
    error = ''
    if request.method == 'POST':
        try:
            numpages, txtfile = extract_from_pdf(request.files['file'], request.form.get('password'))
        except Exception:
            error = 'Check file input and password!'
        else:
            content, matches2 = parse_mpesa_content(txtfile)
            if matches2:
                name = find_name(matches2)
                title = name + ' ' +'MPESA'+'.xlsx'
            if content:
                #content2 = exec_analytics(content)
                #pandas operation
                positives = paidin(content)
                negatives = withdrawal(content)
                dfslist = listing(positives, negatives)
                sheets = ['PAID IN DATA', 'WITHDRAWN DATA']
                output = dfs_tabs(dfslist, sheets, content)
                resp = make_response((output.getvalue(), {
                    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    'Content-Disposition': 'attachment; filename={}.xlsx'.format(title)
                }))
            else:
                resp =  make_response({'response' : content}, {
                    'Content-Type': 'application/json',
                })
            return resp

    return render_template('index.html', error=error)
