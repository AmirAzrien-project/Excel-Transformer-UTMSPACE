Excel Transformer (Laravel)

This project is a Laravel-based Excel processing tool used to clean, transform, and standardize applicant data from Excel files. It supports multiple sheets and applies rule-based transformations to ensure data consistency.

SUPPORTED SHEETS

- PUPW
- UTM-IDP
- Foundation

MAIN FEATURES

- Excel file upload and processing
- Data transformation engine (rule-based)
- Address splitting and normalization
- City abbreviation standardization
- Subject code mapping (legacy to new codes)
- Foreign applicant detection and field handling
- Data validation and anomaly reporting
- Export cleaned Excel output

REQUIREMENTS

- PHP 8.3+
- Composer
- Node.js (for frontend assets)
- MySQL (if database features enabled)

INSTALLATION

1. Clone repository
   git clone <repo-url>

2. Install dependencies
   composer install
   npm install

3. Setup environment
   cp .env.example .env
   php artisan key:generate

4. Run migrations (if needed)
   php artisan migrate

5. Start development server
   php artisan serve

USAGE FLOW

1. Upload Excel file via web interface
2. System processes and applies transformation rules
3. Review validation warnings (if any)
4. Export cleaned Excel file

IMPORTANT NOTES

- Do not upload vendor/ or node_modules/ to repository
- Build files are excluded from version control
- Ensure correct mapping rules are configured in config/excel_rules.php
- Large Excel files may require increased PHP memory limit

PROJECT STRUCTURE (CORE)

- app/Services/ExcelTransformer.php (main transformation logic)
- app/Http/Controllers/ExcelController.php (request handling)
- config/excel_rules.php (mapping rules)
- resources/views (UI)

LICENSE
Internal use only
