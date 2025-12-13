# Python Resume Builder

A professional resume generator that creates beautifully formatted Word documents (.docx) from JSON data files.

## Features

- **Externalized Content**: All resume content is stored in separate JSON files
- **Version Control**: Automatically creates versioned files (_1, _2, _3, etc.)
- **Customizable Layout**: Professional formatting with configurable margins
- **Template System**: Includes template files with placeholder data for GitHub sharing

## Project Structure

```
python-resume-builder/
├── generate-resume.py          # Main script
├── data/                        # Your personal resume data (NOT committed to Git)
│   ├── personal-info.json
│   ├── summary.json
│   ├── skills.json
│   ├── experience.json
│   ├── education.json
│   ├── certifications.json
│   └── section-labels.json
├── data-templates/              # Template files with placeholder data (for GitHub)
│   ├── personal-info.json
│   ├── summary.json
│   ├── skills.json
│   ├── experience.json
│   ├── education.json
│   ├── certifications.json
│   └── section-labels.json
└── generated/                   # Output folder for generated resumes (NOT committed)
```

## Getting Started

### 1. Install Dependencies

```bash
pip3 install python-docx
```

### 2. Set Up Your Data

Copy the template files to create your personal data:

```bash
cp -r data-templates/ data/
```

Then edit the JSON files in the `data/` folder with your information:

- **personal-info.json**: Name, contact details, location
- **summary.json**: Professional summary
- **skills.json**: Skills organized by category
- **experience.json**: Job roles with bullet points
- **education.json**: Degrees and institutions
- **certifications.json**: List of certifications
- **section-labels.json**: Section header labels (optional customization)

### 3. Generate Your Resume

```bash
python3 generate-resume.py
```

Your resume will be created in the `generated/` folder with automatic version numbering.

## Customization

### Page Layout

The script is configured with 0.5-inch margins on all sides. You can adjust these in [generate-resume.py:30-34](generate-resume.py#L30-L34).

### Styling

- Font: Arial, 10pt (configurable in the script)
- Section headers: White text on dark blue background
- Professional color scheme throughout

## Git Usage

The `.gitignore` file ensures your personal data stays private:
- `data/` folder is excluded (your personal information)
- `generated/` folder is excluded (generated resume files)
- `data-templates/` folder IS committed (for sharing the project)

## Future Enhancements

- Frontend UI for easy content management
- Multiple resume templates
- Export to PDF
- Custom styling options

## License

MIT License - feel free to use and modify for your needs.
