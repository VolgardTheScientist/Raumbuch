import streamlit as st
import pandas as pd
import ifcopenshell
import ifcopenshell.util.element
import os
import tempfile
from io import BytesIO

# Load the data from the excel file
excel_file_path = r'O:\01 Laufende Projekte\692 WB Erweiterung Klinik Hirslanden Aarau\02 Planunterlagen\02 BIM AKTUELL\020 Raumbuch\Python\Raumbuch_Attributenliste_IFC.xlsx'
property_pset_pairs = pd.read_excel(excel_file_path, engine='openpyxl', usecols=[0, 1]).astype(str).applymap(str.strip)

st.title("692 Raumbuch IFC to XLS")

ifc_file = st.file_uploader("Upload an IFC file", type="ifc")

if ifc_file is not None:
    uploaded_file_name = ifc_file.name  # Save the name before overwriting the variable

    # Create a temporary file
    tfile = tempfile.NamedTemporaryFile(delete=False) 
    tfile.write(ifc_file.getvalue())
    tfile.close()

    # Open the file using ifcopenshell
    ifc_file = ifcopenshell.open(tfile.name)
    rooms = ifc_file.by_type('IfcSpace')


    # Select all spaces which represent rooms in an IFC file
    rooms = ifc_file.by_type('IfcSpace')

    data = []

    # Iterate over all rooms
    for room in rooms:
        room_data = {"global_id": room.GlobalId, "name": room.Name}

        # Get all property sets of the room
        for relDefinesByProperties in room.IsDefinedBy:
            if relDefinesByProperties.is_a("IfcRelDefinesByProperties"):
                property_set = relDefinesByProperties.RelatingPropertyDefinition

                # Ensure we're working with an IfcPropertySet (there can be IfcElementQuantity as well)
                if property_set.is_a("IfcPropertySet"):
                    # Iterate over properties in the set
                    for property in property_set.HasProperties:
                        # Strip leading and trailing spaces from property and PSet name
                        property_name_stripped = property.Name.strip()
                        pset_name_stripped = property_set.Name.strip()

                        # Combine property name and PSet name into a single key
                        combined_key = f"{pset_name_stripped}__{property_name_stripped}"

                        # Check if the combined key is in the list
                        if property.is_a("IfcPropertySingleValue") and \
                                any((property_pset_pairs['PSet'] + '__' + property_pset_pairs['Property']) == combined_key) \
                                and property.Name != 'NetFloorArea':
                            # Store the property value in the room_data dictionary under the combined key
                            room_data[combined_key] = property.NominalValue.wrappedValue if property.NominalValue else None

        # Append room data to the list
        data.append(room_data)

    # Iterate over all rooms to get BaseQuantities

    quantity_data = []

    for room in rooms:
        room_data = {"global_id": room.GlobalId, "name": room.Name}

        # Get the BaseQuantities
        psets = ifcopenshell.util.element.get_psets(room, qtos_only=True)
        if 'BaseQuantities' in psets and 'NetFloorArea' in psets['BaseQuantities']:
            room_data['BaseQuantities__NetFloorArea'] = psets['BaseQuantities']['NetFloorArea']

        # Append room quantity data to the list
        quantity_data.append(room_data)

    # Convert list of dictionaries to pandas DataFrame
    df = pd.DataFrame(data)

    # Convert list of QTO dictionaries to pandas DataFrame
    dfQto = pd.DataFrame(quantity_data)

    # Create a dictionary to map combined key column names back to Property names only
    column_rename_dict = dict(zip((property_pset_pairs['PSet'] + '__' + property_pset_pairs['Property']).tolist(), property_pset_pairs['Property'].tolist()))

    # Adjust column_rename_dict to account for NetFloorArea
    column_rename_dict['BaseQuantities__NetFloorArea'] = 'Fläche Ist [m²]'

    # Arrange DataFrame columns according to property-PSet combined keys order and then rename columns
    df = df.reindex(columns=['global_id', 'name'] + (property_pset_pairs['PSet'] + '__' + property_pset_pairs['Property']).tolist()).rename(columns=column_rename_dict)

    # Add Qtos to DataFrame with PSets
    df['Fläche Ist [m²]'] = dfQto['BaseQuantities__NetFloorArea']

    # Round the 'Fläche Ist [m²]' column to two decimal places
    df['Fläche Ist [m²]'] = df['Fläche Ist [m²]'].round(2)
    print('Laufnummer und Fläche IST:')
    print(df['Fläche Ist [m²]'])

    # Remove rows where 'name' equals '*'
    df = df.loc[df['name'] != '*']

    # Convert 'name' column to numeric format
    df['name'] = pd.to_numeric(df['name'], errors='coerce')

    # Remove 'global_id' column
    df = df.drop(columns=['global_id'])

    # Rename 'name' column to 'Laufnummer'
    df = df.rename(columns={'name': 'Laufnummer'})

    # Export DataFrame to Excel
    df.to_excel('692_RAU_IFC_to_Python_output_V2.5_2023.06.27.xlsx', index=False)
    print('Das Raumbuch wurde erfolgreich exportiert!')


    # Make a BytesIO object to store the excel data in memory
    output = BytesIO()

    # Convert DataFrame to Excel and use streamlit to download
    excel_writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(excel_writer, index=False)
    excel_writer.save()

    # Convert BytesIO to a streamlit download button
    st.download_button(
        label="Download output XLS file",
        data=output.getvalue(),
        file_name="output.xls",
        mime="application/vnd.ms-excel"
    )

    # Display success message
    st.success(f'Dein Raumbuch aus dem IFC Modell ({uploaded_file_name}) wurde erfolgreich exportiert!')

    st.write('Unten sehen Sie eine Vorschau Ihres Raumbuchs:')
    def rename_cols(df):
        cols = pd.Series(df.columns)
        for dup in cols[cols.duplicated()].unique(): 
            cols[cols[cols == dup].index.values.tolist()] = [dup + '_' + str(i) if i != 0 else dup for i in range(sum(cols == dup))]
        df.columns = cols
        return df

    # create a copy of your dataframe
    df_temp = df.copy()

    # rename the duplicate columns in the copied dataframe
    df_temp = rename_cols(df_temp)

    # now you can display the dataframe with unique column names
    st.dataframe(df_temp)


    # Once all processing is done, remove the temporary file
    os.remove(tfile.name)

