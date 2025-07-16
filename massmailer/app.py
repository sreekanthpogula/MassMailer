from jinja2 import Template
from massmailer.templates import email_draft

template_str = email_draft

template = Template(template_str)
rendered = template.render(associate="Sreekanth Pogula")
print(rendered)