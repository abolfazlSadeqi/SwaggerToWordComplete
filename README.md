# API Documentation Generator

Generate **professional, standardized Word (DOCX) API documentation** directly from a **Swagger / OpenAPI JSON file**.
This tool is designed for developers, technical writers, and companies that need **clean, formatted, business-ready API documentation** without spending hours writing it manually.

---

## 📌 Key Features

### 🔹 1. Generate DOCX From Swagger

Just upload a Swagger/OpenAPI JSON file → the system automatically generates:

* A complete API documentation (DOCX)
* Clean formatting, standardized template
* Automatic table generation
* Full endpoint descriptions
* Request, Response, Headers, Body parameters
* Automatic Type detection

---

### 🔹 2. Automatic Index and TOC (Table of Contents)

* Clickable index
* Method list (URL, Method, Summary, Description)
* Sorted by method or route

---

### 🔹 3. Business & Technical Introduction

Optional sections include:

* **Business Introduction**
* **General Overview**
* **Document Purpose**
* **Shared Models**
* **Shared Errors**
* **Developer notes**

---

### 🔹 4. Change Log Support

You can provide a JSON change-log like:

```json
[
  {
    "version": "1.0.0",
    "date": "2025-11-01",
    "developer": "Test",
    "notes": "Initial API version"
  }
]
```

The tool automatically builds a formatted Change Log table.

---

### 🔹 5. RTL / LTR + Font Customization

Full support for:

* RTL Languages (Persian, Arabic, Hebrew)
* LTR Languages (English, European languages)
* Custom fonts: Tahoma, IRANSans, Calibri, B Nazanin, and more
* Adjustable font size (titles, body, headings)

---

### 🔹 6. Metadata & Settings

You can configure:

* Margins
* Border size
* Colors
* Header sizes
* Show/hide page numbers
* Enable/disable TOC
* Business logo (future update)

---

### 🔹 7. Clean UI With Tabs

The form includes:

* **File Upload**
* **Table Titles**
* **Metadata**
* **Settings**
* **Font & Direction**
* **Advanced Settings**

All fields have default values so you can generate documentation quickly.

---

## ⚙️ How It Works

1. User uploads **Swagger/OpenAPI JSON**
2. User fills optional fields (intro, table titles, font, direction…)
3. System parses the Swagger file and extracts:

   * Endpoints
   * Methods
   * Models
   * Request / Response bodies
   * Parameters
   * Authentication requirements
4. A fully-formatted **Word document** is generated using a consistent template.
5. User downloads the generated DOCX.

---

## 🚀 Usage

1. Open the app in browser:

```
http://localhost:5000
```

2. Go to **API Documentation Generator**
3. Upload your Swagger JSON file
4. (Optional) Fill metadata, titles, font, direction
5. Click **Generate DOCX**
6. Download your generated API documentation


Just tell me **“add badges”** or **“add Persian version”**.
