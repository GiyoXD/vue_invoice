# run_test.ps1
# Simple script to run the Snitch Chain Tracing test

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "   Starting Invoice Pipeline Test Run" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

# Run the pipeline runner module
# Target: JF.xlsx
# Output: output/generated_invoices
python -m core.invoice_workflow "temp_uploads\JF25057.xlsx" --output "output/generated_invoices"

Write-Host ""
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "            Test Run Finished" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "Check output/generated_invoices for results."
Write-Host "Check run_log/ for snitch traces."
Write-Host ""

Read-Host "Press Enter to exit..."
