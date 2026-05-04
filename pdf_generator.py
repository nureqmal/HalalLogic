from jinja2 import Template
import pdfkit

config = pdfkit.configuration(
    wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
)

def generate_pdf(company, policy, materials, filename="manual.pdf"):

    rows = ""
    for m in materials:
        rows += f"<tr><td>{m[0]}</td><td>{m[1]}</td><td>{m[2]}</td><td>{m[3]}</td></tr>"

    html = Template("""
    <html>
    <head>
    <style>
    body { font-family: Arial; margin:40px;}
    h1 {text-align:center;}
    h2 {border-bottom:2px solid black;}
    table {width:100%; border-collapse: collapse;}
    th, td {border:1px solid black; padding:8px;}
    </style>
    </head>

    <body>
    <h1>IHCS MANUAL</h1>

    <h2>Company Profile</h2>
    <p><b>Name:</b> {{company[0]}}</p>
    <p><b>Address:</b> {{company[1]}}</p>
    <p><b>Contact:</b> {{company[2]}}</p>

    <h2>Halal Policy</h2>
    <p>{{policy}}</p>

    <h2>Raw Materials</h2>
    <table>
    <tr><th>Name</th><th>Supplier</th><th>Halal Cert</th><th>Expiry</th></tr>
    """ + rows + """
    </table>

    </body>
    </html>
    """).render(company=company, policy=policy)

    pdfkit.from_string(html, filename, configuration=config)
    return filename
