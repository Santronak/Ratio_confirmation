import streamlit as st
import pandas as pd
import os
import glob
import tempfile
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as L
import io

# ================= SETTINGS =================
def process_excel(file_content, service, ratio):
    """Process the Excel file with given service and ratio"""
    
    # Create temporary files
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_input:
        tmp_input.write(file_content)
        tmp_input_path = tmp_input.name
    
    try:
        # ================= SETTINGS =================
        service = service.upper().strip()
        if service not in ("FPS", "IRIS"):
            raise Exception("Service must be FPS or IRIS")
        
        if isinstance(ratio, int):
            ratios = [ratio]
        elif isinstance(ratio, str) and "-" in ratio:
            a, b = map(int, ratio.split("-"))
            ratios = list(range(a, b + 1, 5))
        else:
            ratios = [int(ratio)]
        
        # ================= FILE =================
        folder = os.path.dirname(tmp_input_path)
        
        # Check if there are any Excel files in the folder
        files = [tmp_input_path]  # Use the uploaded file directly
        
        infile = files[0]
        outfile = os.path.splitext(infile)[0] + "_Output.xlsx"
        
        wb = load_workbook(infile, keep_vba=infile.lower().endswith(".xlsm"))
        
        # ================= SHEETS =================
        if "Summary" in wb.sheetnames:
            del wb["Summary"]
        
        ws = wb[wb.sheetnames[0]]
        ws.title = str(ratios[0])
        
        for x in ratios[1:]:
            wb.copy_worksheet(ws).title = str(x)
        
        # ================= PROCESS =================
        months = ["Jan'", "Feb'", "Mar'", "Apr'", "May'", "Jun'",
                  "Jul'", "Aug'", "Sep'", "Oct'", "Nov'", "Dec'"]
        
        for rv in ratios:
            ws = wb[str(rv)]
            div = rv + 1
            
            # Month columns
            mcols = [
                c for c in range(1, ws.max_column + 1)
                if any(m in str(ws.cell(1, c).value or "") for m in months)
            ]
            
            if not mcols:
                raise Exception(f"No month columns in sheet {rv}")
            
            # Total / Max Candidate
            total = ws.max_column + 1
            maxcand = total + 1
            ws.cell(1, total).value = "Total Candidate"
            ws.cell(1, maxcand).value = "Max Candidate"
            
            for r in range(2, ws.max_row + 1):
                refs = ",".join(f"{L(c)}{r}" for c in mcols)
                ws.cell(r, total).value = f"=SUM({refs})"
                ws.cell(r, maxcand).value = f"=MAX({refs})"
            
            # Opr columns
            opr_start = maxcand + 1
            opr = []
            
            for i, c in enumerate(mcols):
                oc = opr_start + i
                opr.append(oc)
                ws.cell(1, oc).value = f"{ws.cell(1,c).value} Opr"
                
                for r in range(2, ws.max_row + 1):
                    ws.cell(r, oc).value = f"=ROUNDUP({L(c)}{r}/{div},0)"
            
            # Max Opr
            maxopr = opr_start + len(opr)
            ws.cell(1, maxopr).value = "Max Opr"
            
            for r in range(2, ws.max_row + 1):
                refs = ",".join(f"{L(c)}{r}" for c in opr)
                ws.cell(r, maxopr).value = f"=MAX({refs})"
            
            # Service columns
            names = (
                ["Tab", "FPS", "OTG", "Hologram", "Id Card", "Jacket"]
                if service == "FPS"
                else ["Tab", "IRIS", "FPS", "OTG", "Hologram", "Id Card", "Jacket"]
            )
            
            sc = {}
            for i, name in enumerate(names):
                sc[name] = maxopr + 1 + i
                ws.cell(1, sc[name]).value = name
            
            T, M = L(total), L(maxopr)
            tab = L(sc["Tab"])
            
            for r in range(2, ws.max_row + 1):
                ws.cell(r, sc["Tab"]).value = (
                    f'=IF({M}{r}=0,0,IF({M}{r}>=8,'
                    f'IF({M}{r}<16,{M}{r}+2,{M}{r}+3),{M}{r}+1))'
                )
                
                if service == "FPS":
                    ws.cell(r, sc["FPS"]).value = f"={tab}{r}"
                else:
                    ws.cell(r, sc["IRIS"]).value = f"={tab}{r}"
                    ws.cell(r, sc["FPS"]).value = "=1"
                
                ws.cell(r, sc["OTG"]).value = f"={tab}{r}*2"
                ws.cell(r, sc["Hologram"]).value = f"=ROUNDUP({T}{r}/100,0)+1"
                ws.cell(r, sc["Id Card"]).value = f"={M}{r}+1"
                ws.cell(r, sc["Jacket"]).value = f"={M}{r}"
        
        # ================= SUMMARY =================
        summary = wb.create_sheet("Summary", 0)
        first = wb[str(ratios[0])]
        
        headers = [str(first.cell(1,c).value or "").strip()
                   for c in range(1, first.max_column + 1)]
        
        hm = {h:i+1 for i,h in enumerate(headers) if h}
        
        opr_headers = [
            h for h in headers
            if h.endswith(" Opr") and h != "Max Opr"
        ]
        
        service_names = (
            ["Tab","FPS","OTG","Hologram","Id Card","Jacket"]
            if service == "FPS"
            else ["Tab","IRIS","FPS","OTG","Hologram","Id Card","Jacket"]
        )
        
        summary_headers = (
            ["Ratio","Total Center","Total Candidate","Max Candidate"]
            + opr_headers
            + ["Max Opr"]
            + service_names
            + ["Avg"]
        )
        
        for c,h in enumerate(summary_headers,1):
            summary.cell(1,c).value = h
        
        for r, rv in enumerate(ratios, 2):
            sh = f"'{rv}'"
            summary.cell(r,1).value = rv
            
            source = {
                "Total Center": "Total Candidate",
                "Total Candidate": "Total Candidate",
                "Max Candidate": "Max Candidate",
                "Max Opr": "Max Opr",
                **{x:x for x in opr_headers},
                **{x:x for x in service_names}
            }
            
            for c,h in enumerate(summary_headers[1:-1],2):
                src = hm[source[h]]
                fn = "COUNT" if h == "Total Center" else "SUM"
                summary.cell(r,c).value = f"={fn}({sh}!{L(src)}:{L(src)})"
            
            # Avg = Max Candidate / Max Opr
            mc = summary_headers.index("Max Candidate") + 1
            mo = summary_headers.index("Max Opr") + 1
            av = summary_headers.index("Avg") + 1
            
            summary.cell(r,av).value = (
                f"=IFERROR(ROUNDUP({L(mc)}{r}/{L(mo)}{r},0),0)"
            )
        
        # ================= FORMAT / SAVE =================
        summary.freeze_panes = "A2"
        
        for col in summary.columns:
            letter = L(col[0].column)
            summary.column_dimensions[letter].width = min(
                max(len(str(x.value or "")) for x in col) + 2, 8
            )
        
        try:
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
            wb.calculation.calcMode = "auto"
        except:
            pass
        
        # Save to bytes
        output = io.BytesIO()
        wb.save(output)
        wb.close()
        output.seek(0)
        
        return output, ratios, service
        
    finally:
        # Clean up temporary file
        try:
            os.unlink(tmp_input_path)
        except:
            pass

# ================= STREAMLIT UI =================
st.set_page_config(
    page_title="Excel Processor",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel File Processor")
st.markdown("Process Excel files for FPS and IRIS services with ratio calculations")

# Sidebar
with st.sidebar:
    st.header("⚙️ Settings")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Upload Excel File",
        type=['xlsx', 'xlsm', 'xls'],
        help="Upload an Excel file (.xlsx, .xlsm, .xls)"
    )
    
    # Service selection
    service = st.selectbox(
        "Select Service",
        options=["FPS", "IRIS"],
        help="Choose the service type (FPS or IRIS)"
    )
    
    # Ratio input
    ratio_input = st.text_input(
        "Enter Ratio",
        value="45-60",
        help="Enter a single number (e.g., 75) or a range (e.g., 20-60)"
    )
    
    # Process button
    process_button = st.button(
        "🚀 Process File",
        type="primary",
        use_container_width=True
    )

# Main content area
col1, col2 = st.columns([2, 1])

with col1:
    if uploaded_file:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        st.info(f"📄 File size: {uploaded_file.size / 1024:.2f} KB")
        
        # Preview uploaded file
        try:
            df = pd.read_excel(uploaded_file, sheet_name=0, nrows=5)
            st.subheader("📋 File Preview")
            st.dataframe(df, use_container_width=True)
            uploaded_file.seek(0)  # Reset file pointer
        except Exception as e:
            st.warning(f"Could not preview file: {str(e)}")
    else:
        st.info("📤 Please upload an Excel file to begin")
        st.markdown("""
        ### Instructions:
        1. Upload an Excel file (.xlsx, .xlsm, or .xls)
        2. Select the service (FPS or IRIS)
        3. Enter the ratio (single value or range)
        4. Click "Process File" to generate output
        5. Download the processed file
        """)

with col2:
    if uploaded_file and process_button:
        with st.spinner("⏳ Processing file..."):
            try:
                # Read file content
                file_content = uploaded_file.read()
                
                # Process the file
                output_file, ratios, service_name = process_excel(
                    file_content, service, ratio_input
                )
                
                st.success("✅ File processed successfully!")
                
                # Display summary
                st.subheader("📊 Processing Summary")
                st.info(f"**Service:** {service_name}")
                st.info(f"**Ratios:** {', '.join(map(str, ratios))}")
                
                # Download button
                st.download_button(
                    label="📥 Download Output File",
                    data=output_file,
                    file_name=f"{os.path.splitext(uploaded_file.name)[0]}_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"❌ Error processing file: {str(e)}")
                st.exception(e)

# Footer
st.markdown("---")
st.markdown("Built with ❤️ using Streamlit")

# Additional information
with st.expander("ℹ️ How it works"):
    st.markdown("""
    ### Process Workflow:
    
    1. **File Upload**: Upload your Excel file containing data
    2. **Service Selection**: Choose between FPS or IRIS service
    3. **Ratio Configuration**: 
       - Single value: e.g., `75`
       - Range: e.g., `20-60` (generates values at 5-interval steps)
    4. **Processing**: The system will:
       - Create ratio-specific sheets
       - Calculate Total and Max Candidate values
       - Generate Opr columns with ROUNDUP formulas
       - Create service-specific columns (Tab, FPS/IRIS, OTG, etc.)
       - Generate a Summary sheet with all calculations
    5. **Download**: Get the processed file with all formulas and calculations
    """)
