import streamlit as st
import pandas as pd
import io

# --- Main Streamlit App ---

st.set_page_config(page_title="Excel File Processor", layout="centered")

st.title("Excel File Processor")
st.write(
    "Upload your Excel file to group data and consolidate columns. "
    "The script will group rows by 'Composition', 'Temperature', 'Orientation', 'Stress (MPa)', and 'Test duration', "
    "and then merge the 'Strain (%)' and 'Time' columns into comma-separated lists."
)


# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    st.success(f"File '{uploaded_file.name}' uploaded successfully!")

    if st.button("Process File"):
        try:
            with st.spinner("Processing your file..."):
                # Read the uploaded Excel file
                df = pd.read_excel(uploaded_file)

                # Clean up column names by stripping leading/trailing whitespace
                df.columns = df.columns.str.strip()
                
                # Define the columns to group by
                grouping_columns = [
                    'Composition', 'Temperature', 'Orientation',
                    'Stress (MPa)', 'Test duration'
                ]

                # Check for missing grouping columns
                missing_cols = [col for col in grouping_columns if col not in df.columns]
                if missing_cols:
                    st.error(f"The uploaded file is missing the following required columns for grouping: {', '.join(missing_cols)}")
                else:
                    # Define the columns to aggregate
                    aggregation_defs = {
                        'Strain (%)': lambda x: ','.join(x.astype(str)),
                        'Time': lambda x: ','.join(x.astype(str))
                    }

                    # Check for missing aggregation columns
                    agg_cols = list(aggregation_defs.keys())
                    missing_agg_cols = [col for col in agg_cols if col not in df.columns]
                    if missing_agg_cols:
                         st.error(f"The uploaded file is missing the following required columns for aggregation: {', '.join(missing_agg_cols)}")
                    else:
                        # Perform the grouping and aggregation
                        processed_df = df.groupby(grouping_columns).agg(aggregation_defs).reset_index()

                        # Add the empty 'UTS' column if it doesn't exist
                        if 'UTS' not in processed_df.columns:
                            processed_df['UTS'] = pd.NA

                        # Reorder columns to match the desired output layout
                        output_columns_order = [
                            'Composition', 'Temperature', 'Orientation', 'Stress (MPa)', 
                            'UTS', 'Test duration', 'Strain (%)', 'Time'
                        ]
                        
                        # Ensure all columns are present before reordering
                        final_columns = [col for col in output_columns_order if col in processed_df.columns]
                        processed_df = processed_df[final_columns]

                        st.success("File processed successfully!")
                        
                        # Convert DataFrame to an in-memory Excel file
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            processed_df.to_excel(writer, index=False, sheet_name='Processed_Data')
                        
                        st.download_button(
                            label="Download Processed File",
                            data=output.getvalue(),
                            file_name="processed_data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )


        except Exception as e:
            st.error(f"An unexpected error occurred: {e}")
else:
    st.info("Please upload an Excel file to begin.") 