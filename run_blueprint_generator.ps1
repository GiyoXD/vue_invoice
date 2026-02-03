
# Run the Blueprint Generator for JF25060 with custom client name JF_TEST
# Input: temp_uploads/CT&INV&PL JF25060 FCA .xlsx
# Output Client: JF_TEST

python -m core.blueprint_generator.blueprint_generator "temp_uploads\CT&INV&PL JF25060 FCA .xlsx" --prefix "JF_TEST" -v

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Blueprint Generation Completed for JF_TEST" -ForegroundColor Cyan
Write-Host "Check database/blueprints/bundled/JF_TEST for output" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
