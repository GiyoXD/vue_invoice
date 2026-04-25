







### DeepSheet
- this is the sheet to source the invoice no from 
invoice_no: A2,
invoice_ref: B2,
invoice_date: C2
total_net: D2,
total_gross: E2

# todo
add safety measure for the cbm, and net weight placement, some sheet are freaky by the creator, cause the formula not working as expected

# new task
- track the invoiec gen for the exeption bug
- grand total bug
- refactor file orgininzation


does it look good for you? too many argument
class GenerateRequest(BaseModel):
    identifier: str
    json_path: str
    invoice_no: str
    invoice_date: str
    invoice_ref: Optional[str] = ""
    generate_standard: bool = True
    generate_custom: bool = False
    generate_daf: bool = False
    generate_kh: bool = False
    generate_vn: bool = False
    price_adjustment: Optional[List[List[Any]]] = None
    global_unit_price: Optional[float] = None  # For 'net' pricing mode (shipping lists)
    pricing_net_weight: bool = False
    auto_fit: bool = True