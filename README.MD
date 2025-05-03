# Scripts 

Making my life easier, one script at a time.

## CC Analyzer

### Prerequisites

- Head over to `python` folder
- Enable Google Sheets and Google Drive API in Google Cloud Console
- Share targeted Google sheet file to Service account email as "Editor"
- Add Service Account's JSON file as `credentials.json` under `src` folder
- Download PDF file of your CC bill

### Install Dependencies

```
pip install -r requirements.txt
```

### Running the app

```
python3 src/main.py --sheet-name "NAME_OF_GOOGLE_SHEET" --billing-period "MMMM YYYYY" --pdf-path "pdf/path" --pdf-password "pdf_password" 
```

> Supported Credit Card: Security Bank

#### Get in touch

<p>
<a href="https://twitter.com/crrmacarse">
<img src="https://img.shields.io/badge/crrmacarse%20-%231DA1F2.svg?&style=for-the-badge&logo=Twitter&logoColor=white"/>
</a>
<a href="https://www.linkedin.com/in/christian-ryan-r-macarse-692974166/">
<img src="https://img.shields.io/badge/linkedin%20-%230077B5.svg?&style=for-the-badge&logo=linkedin&logoColor=white"/>
</a>
<a href="https://stackoverflow.com/users/10030210/crrmacarse?tab=profile">
<img src="https://cdn.sstatic.net/Sites/stackoverflow/company/Img/logos/so/so-logo.svg?v=a010291124bf" width="110px" />
</a>
</p>
