# 📃 SQDoc

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)


Welcome to the manual of SQDoc - a script to create technical documentation of SQL database properties.

## ⚙️ Main functionalities
___
Script performs details gathering of a Microsoft SQL database and formats obtained details as a ``*.docx`` document.
Information scope can be configured and consist of:
- database properties,
- database tables details,
- database stored procedures details.

## ➕ Dependencies
___
To run, script requires an SQL database and installed python-docx library:

``pip install python-docx``

## 📃 Outputs
___
Script default outputs are:
- logs: all actions logged locally in a ``PATH_LOGS`` configurable location - set in a ``main`` file.
- database detail documents: saved at ``PATH_DOCS`` configurable location - set in a ``main`` file.

## 📝 Annotations
___
To configure document's first page provide details in ``main`` file > ``DOC_PROPERTIES`` variable as follows:
``` python
DOC_PROPERTIES = [
    ['Owner:', 'Owner unit'],
    ['Author:', 'Author Name'],
    ['E-mail:', 'dbadmin@domain.com'],
    ['Version:', '1.0'],
    ['Status:', 'Final'],
    ['Created on:', date.today().strftime("%Y-%m-%d")]
]
```

To configure document scope provide details in ``main`` file > ``DOC_CONTENT`` variable as follows:

```python
DOC_CONTENT = {
    'db configuration': True,
    'db tables': True,
    'db procedures': True
}
```

## 📜 License
___
Project is available under terms of **[Apache License 2.0](http://www.apache.org/licenses/LICENSE-2.0)**.  
Full license text can be found in file: [LICENSE](./LICENSE).

---

© 2025 **Vismaanen** — simple coding for simple life