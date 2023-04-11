import pandas as pd
import datetime
import openpyxl
import os
import xlrd

pd.options.mode.chained_assignment = None

print('''
Please type the number for the supplier you need data formatted for:
1: A
2: B
3: C
4: D
5: E
''')

supplier = input('> ')

if int(supplier) == 1: # A

    path = input('''Enter the path to the file including the extension e.g. xlsx (Note: you can right-click on the file and select 'Copy as path') ''')
    # Removing speechmarks from path if present
    path = path.lstrip('\"')
    path = path.rstrip('\"')
    dirname = os.path.dirname(path)

    print('Importing file...')

    data = pd.read_excel(path, sheet_name=None)
    df = pd.read_excel(path)
    sheet_names = list(data.keys())

    df = df.fillna("")      # Fill NaN vales
    for index,row in df.iterrows():     # Find where data starts
        if row[0]:
            start = index
            break

    if len(sheet_names) == 1:
        df.columns = df.loc[start]      # Set columns as this row
        df = df.iloc[start+1:]      # Set this as new start point of dataframe
        df.reset_index(inplace=True, drop=True)

    elif len(sheet_names) == 2:
        # Merge together
        df1 = data[sheet_names[0]]
        df1.columns = df1.loc[start]    # Set columns as this row
        df1 = df1.iloc[start+1:]      # Set this as new start point of dataframe
        df1.reset_index(inplace=True, drop=True)

        df2 = data[sheet_names[1]]
        df2.columns = df2.loc[start]      # Set columns as this row
        df2 = df2.iloc[start+1:]      # Set this as new start point of dataframe
        df2.reset_index(inplace=True, drop=True)

        df = pd.merge(df1, df2, left_index=True, right_index=True, suffixes=('', '_x'))
        cols_to_drop = [col + '_x' for col in df2.columns if col in df1.columns]
        df.drop(cols_to_drop, axis=1, inplace=True)
        df.reset_index(inplace=True, drop=True)

    else:
        input("Expected either one or two excel sheets in the file, more are present")

    print('File imported. Formatting data...')

    # If 'SKU' is present use that
    # If not, check which columns needed then add together to make 'SKU'

    if 'SKU' in df.columns:
        pass
    else:
        for col in df.columns:
            if 'Item Nbr' in col:
                df[col] = df[col].map(str)

        try:
            # If variety column is present, include
            if 'Variety' in df.columns:
                try:
                    df['SKU'] = df.filter(like='Item Nbr', axis=1).fillna('').agg(''.join,axis=1) + ' ' + df['Prime Item Desc'] + ' ' + df['Variety'] + ' ' + df.filter(like='Size Desc', axis=1).fillna('').agg(''.join,axis=1)
                except:
                    df['SKU'] = df.filter(like='Item Nbr', axis=1).fillna('').agg(''.join,axis=1) + ' ' + df['Shelf Description'] + ' ' + df['Variety'] + ' ' + df.filter(like='Size Desc', axis=1).fillna('').agg(''.join,axis=1)
            else:
                try:
                    df['SKU'] = df.filter(like='Item Nbr', axis=1).fillna('').agg(''.join,axis=1) + ' ' + df['Prime Item Desc'] + ' ' + df.filter(like='Size Desc', axis=1).fillna('').agg(''.join,axis=1)
                except:
                    df['SKU'] = df.filter(like='Item Nbr', axis=1).fillna('').agg(''.join,axis=1) + ' ' + df['Shelf Description'] + ' ' + df.filter(like='Size Desc', axis=1).fillna('').agg(''.join,axis=1)
        except:
            print('''SKU Name unable to be extracted from columns, please ensure "Prime Item Nbr", "Prime Item Desc", 
            and "Prime Size Desc" columns are included in data or rename columns where needed''')
        
    df = df[['SKU'] + [col for col in df.columns if col != 'SKU']]    # Move SKU column to first

    # If EPOS in column names with week then adapt data to make new week column
    if True in df.columns.str.contains('EPOS'):
        # Get the column names
        cols = df.columns.tolist()

        # Get the weeks
        weeks = [col.split(" ")[0] for col in cols if col.endswith("EPOS Sales")]

        # Get the values and units columns
        values_cols = [col for col in cols if col.endswith("EPOS Sales")]
        units_cols = [col for col in cols if col.endswith("EPOS Qty")]

        df["Week"] = ""

        # Create a new DataFrame to store the modified data
        result = pd.DataFrame(columns=["SKU", "Store Nbr", "£ Value", "Units", "Week"])

        # Fill in the values and units columns
        for week, values_col, units_col in zip(weeks, values_cols, units_cols):
            temp_df = df[["SKU", "Store Nbr", values_col, units_col]]
            temp_df.rename(columns={values_col: "£ Value", units_col: "Units"}, inplace=True)
            temp_df["Week"] = week
            result = pd.concat([result, temp_df])
            result.reset_index(inplace=True, drop=True)

    else: # If separate EPOS Sales and EPOS Qty rows
        for column in df:   # Find which column contains rows showing EPOS Sales + Qty
            try:
                if df[column].str.contains('EPOS').any() == True:
                    break
                else:
                    pass
            except:
                pass

        idx = df.columns.get_loc(column)

        # Transpose dataframe based on columns
        df1 = df.melt(id_vars=['SKU', 'Store Nbr', 'Data Type'],
                    value_vars=df.columns[idx+1:],
                    var_name='Week', 
                    value_name='Value')

        # Formatting the columns
        df1['Data Type'] = df1['Data Type'].replace({'EPOS Sales': '£ Value', 'EPOS Qty': 'Units'})

        result = df1.pivot_table(index=['SKU', 'Store Nbr', 'Week'], 
                                columns='Data Type', 
                                values='Value').reset_index()


    print('Exporting file... (may take a few minutes depending on size of file)')

    # Excel has a size limit, if dataframe over this size then split into different excel tabs
    if len(result) < 900000:
        result.to_excel(str(dirname) + '\\AFormatted.xlsx', index=False)
    else:
        # Split into multiple tabs then export, file too big for excel single sheet
        max_sheet = 900000
        num_sheet = (len(result) // max_sheet) + 1

        writer = pd.ExcelWriter(str(dirname) + '\\AFormatted.xlsx')

        for i in range(num_sheet):
            start = i * (len(result) // num_sheet)
            end = (i + 1) * (len(result) // num_sheet)
            chunk = result.iloc[start:end]
            chunk.to_excel(writer, sheet_name='Sheet_{}'.format(i))

        writer.close()

    input('File exported successfully as "AFormatted.xlsx" in the same folder. Press any key to quit.')

elif int(supplier) == 2: # B

    type = input(''' Is this a regular single file formatting (1) or a multiple file formatting (2)? ''')

    if int(type) == 1:
        # SINGLE FILE CODE
        path = input('''Enter the path to the file including the extension e.g. xlsx (Note: you can right-click on the file and select 'Copy as path' and paste it here) ''')
        # Removing speechmarks from path if present
        path = path.lstrip('\"')
        path = path.rstrip('\"')

        dirname = os.path.dirname(path)

        print('Importing file...')

        # Might not be on first sheet
        dfname = pd.ExcelFile(path)
        df = pd.read_excel(path)
        for sheets in dfname.sheet_names[1:]:
            dfnew = pd.read_excel(path, sheet_name=sheets)
            df = pd.concat([df, dfnew])
        df.reset_index(inplace=True, drop=True)

        print('File imported. Formatting data...')

        # Seeing where data starts
        start = int(df.loc[df.iloc[:,0] == 'Geography',:].index[0])

        df.loc[start] = df.loc[start].fillna(method='ffill')

        # Sort out merged cells
        df = df.fillna("")      # Fill NaN vales

        # Change strings to EPOS Sales and EPOS Qty
        df.iloc[start+1] = df.iloc[start+1].apply(lambda x: 'EPOS Sales' if 'Amount' in str(x) else x)
        df.iloc[start+1] = df.iloc[start+1].apply(lambda x: 'EPOS Qty' if 'Volume' in str(x) else x)

        # Organise columns
        df.columns = df.loc[start]  # Set columns as this row
        df = df.iloc[start+1:]      # Set this as new start point of dataframe
        df.reset_index(inplace=True, drop=True)

        # Adding EPOS row to column names
        df.columns = [col + ' ' + df.iloc[0][i] for i, col in enumerate(df.columns)]
        # Removing trailing spaces
        df.columns = [col.rstrip() for col in df.columns]

        df = df.iloc[1:].reset_index(drop=True)

        # Add Store Nbr column
        df.insert(0, 'Store Nbr', df['Geography'].apply(lambda x: x.split('-')[0].strip()))


        # Transposing
        # Get the column names
        cols = df.columns.tolist()

        # Get the weeks
        weeks = [col.split(" ")[1] for col in cols if col.endswith("EPOS Sales")]

        # Get the values and units columns
        values_cols = [col for col in cols if col.endswith("EPOS Sales")]
        units_cols = [col for col in cols if col.endswith("EPOS Qty")]

        df["Week"] = ""

        # Create a new DataFrame to store the modified data
        result = pd.DataFrame(columns=["Product", "Store Nbr", "£ Value", "Units", "Week"])

        # Fill in the values and units columns
        for week, values_col, units_col in zip(weeks, values_cols, units_cols):
            temp_df = df[["Product", "Store Nbr", values_col, units_col]]
            temp_df.rename(columns={values_col: "£ Value", units_col: "Units"}, inplace=True)
            temp_df["Week"] = week
            result = pd.concat([result, temp_df])
            result.reset_index(inplace=True, drop=True)

    elif int(type) == 2:
        # MULTIPLE FILES CODE
        print("Enter the path to *any* of the files in a folder containing the multiple files to format (Note: you can right-click on the file and select 'Copy as path' and paste it here)")
        path = input('''NOTE: THE FOLDER MUST ONLY CONTAIN THE INPUT FILES (can be .xlsx or .csv) ''')
        # Removing speechmarks from path if present
        path = path.lstrip('\"')
        path = path.rstrip('\"')

        dirname = os.path.dirname(path)

        print('Importing file...')

        data_frames = []

        for filename in os.listdir(dirname):
            if filename.endswith(".xlsx"):
                file_path = dirname + '/' + filename
                single_df = pd.read_excel(file_path)
                data_frames.append(single_df)
            elif filename.endswith(".csv"):
                file_path = dirname + '/' + filename
                single_df = pd.read_csv(file_path, sep=';')
                data_frames.append(single_df)
            else:
                print('Invalid file types')

        print('File imported. Formatting data...')

        df = pd.concat(data_frames).reset_index(drop=True)

        df.insert(loc=0, column='SKU', value=df.iloc[:, 0].astype(str) + ' ' + df.iloc[:, 1].astype(str))

        df.drop(df.columns[[1,2,4,5,6]], axis=1, inplace=True)

        cols_to_drop = []
        for col in df.columns[3:]:
            if not (col == 'Sales Value TY' or col == 'Sales Volume TY'):
                cols_to_drop.append(col)

        df = df.drop(cols_to_drop, axis=1)

        result = df

    else:
        input('Invalid option, please select 1 (single file) or 2 (multiple files)')


    print('Exporting file... (may take a few minutes depending on size of file)')

    # Exporting
    if len(result) < 900000:
        result.to_excel(str(dirname) + '\\BFormatted.xlsx', index=False)
    else:
        # Split into multiple tabs then export, file too big for excel single sheet
        max_sheet = 900000
        num_sheet = (len(result) // max_sheet) + 1
        num_sheet

        writer = pd.ExcelWriter(str(dirname) + '\\BFormatted.xlsx')

        for i in range(num_sheet):
            start = i * (len(result) // num_sheet)
            end = (i + 1) * (len(result) // num_sheet)
            chunk = result.iloc[start:end]
            chunk.to_excel(writer, sheet_name='Sheet_{}'.format(i))

        writer.close()

    input('File exported successfully as "BFormatted.xlsx" in the same folder. Press any key to quit.')

elif int(supplier) == 3: # C
    print('Please REMOVE all other tabs in the excel file apart from the one containing sales data!')
    path = input('''Enter the path to the file including the extension e.g. xlsx (Note: you can right-click on the file and select 'Copy as path' and paste it here) ''')
    # Removing speechmarks from path if present
    path = path.lstrip('\"')
    path = path.rstrip('\"')
    dirname = os.path.dirname(path)

    print('Importing file...')

    df = pd.read_excel(path)

    print('File imported. Formatting data...')

    df = df.iloc[:, :10]
    df.drop(df.columns[[1,2,3,5,6,7]], axis=1, inplace=True)
    result = df

    print('Exporting file... (may take a few minutes depending on size of file)')

    # Excel has a size limit, if dataframe over this size then split into different excel tabs
    if len(result) < 900000:
        result.to_excel(str(dirname) + '\\CFormatted.xlsx', index=False)
    else:
        # Split into multiple tabs then export, file too big for excel single sheet
        max_sheet = 900000
        num_sheet = (len(result) // max_sheet) + 1
        num_sheet

        writer = pd.ExcelWriter(str(dirname) + '\\CFormatted.xlsx')

        for i in range(num_sheet):
            start = i * (len(result) // num_sheet)
            end = (i + 1) * (len(result) // num_sheet)
            chunk = result.iloc[start:end]
            chunk.to_excel(writer, sheet_name='Sheet_{}'.format(i))

        writer.close()

    input('File exported successfully as "CFormatted.xlsx" in the same folder. Press any key to quit.')

elif int(supplier) == 4: # D

    path = input('''Enter the path to the D file including the extension e.g. xlsx (Note: you can right-click on the file and select 'Copy as path' and paste it here) ''')
    # Removing speechmarks from path if present
    path = path.lstrip('\"')
    path = path.rstrip('\"')
    dirname = os.path.dirname(path)

    print('Importing file...')

    reference = pd.read_excel(path, 0)

    # Finding where data starts, setting new columns
    columns = reference[reference.eq('Index').any(axis=1, bool_only=True)].index[0]
    reference.columns = reference.iloc[columns]
    reference = reference.drop(columns)

    reference = reference.loc[:, ['Index', 'Measures', 'Product']]
    reference.dropna(subset=['Product'], inplace=True)
    reference = reference.reset_index()

    print('File imported. Formatting data...')

    data = pd.DataFrame()

    for index, row in reference.iterrows():
        sheet_name = str(reference.iloc[index,1])
        df = pd.read_excel(path, sheet_name=sheet_name)

        # Finding the data
        start = int(df.loc[df.iloc[:,0] == 'Geography',:].index[0])
        df.columns = df.iloc[start,:]
        df = df.iloc[start+1:,:]
        df = df.reset_index()

        # Transposing the data
        df = df.melt(id_vars=['Geography'],
                value_vars=df.columns[1:],
                var_name='Week', 
                value_name='Value')
        
        df['SKU'] = reference.iloc[index,3]
        df['Type'] = reference.iloc[index,2]

        cols = df.columns.tolist()
        cols = cols[-2:] + cols[:-2]
        df = df[cols]

        data = pd.concat([data, df])

    result = data.pivot_table(index=['SKU', 'Geography', 'Week'], 
                                columns='Type', 
                                values='Value').reset_index()

    print('Exporting file... (may take a few minutes depending on size of file)')

    # Excel has a size limit, if dataframe over this size then split into different excel tabs
    if len(result) < 900000:
        result.to_excel(str(dirname) + '\\DFormatted.xlsx', index=False)
    else:
        # Split into multiple tabs then export, file too big for excel single sheet
        max_sheet = 900000
        num_sheet = (len(result) // max_sheet) + 1
        num_sheet

        writer = pd.ExcelWriter(str(dirname) + '\\DFormatted.xlsx')

        for i in range(num_sheet):
            start = i * (len(result) // num_sheet)
            end = (i + 1) * (len(result) // num_sheet)
            chunk = result.iloc[start:end]
            chunk.to_excel(writer, sheet_name='Sheet_{}'.format(i))

        writer.close()

    input('File exported successfully as "DFormatted.xlsx" in the same folder. Press any key to quit.')

elif int(supplier) == 5: # E
    path = input('''Enter the path to the file including the extension e.g. xlsx (Note: you can right-click on the file and select 'Copy as path') ''')
    # Removing speechmarks from path if present
    path = path.lstrip('\"')
    path = path.rstrip('\"')

    dirname = os.path.dirname(path)

    print('Importing file...')

    df = pd.read_excel(path)

    print('File imported. Formatting data...')

    # Make SKU Column
    df['SKU'] = df['Item'].astype(str) + ' ' + df['Item Name']
    df = df[['SKU'] + [col for col in df.columns if col != 'SKU']]    # Move SKU column to first

    result = df.loc[:, ['SKU', 'Store', 'Sales £', 'Units', 'Week Number']]

    print('Exporting file... (may take a few minutes depending on size of file)')

    # Excel has a size limit, if dataframe over this size then split into different excel tabs
    if len(result) < 900000:
        result.to_excel(str(dirname) + '\\EFormatted.xlsx', index=False)
    else:
        # Split into multiple tabs then export, file too big for excel single sheet
        max_sheet = 900000
        num_sheet = (len(result) // max_sheet) + 1
        num_sheet

        writer = pd.ExcelWriter(str(dirname) + '\\EFormatted.xlsx')

        for i in range(num_sheet):
            start = i * (len(result) // num_sheet)
            end = (i + 1) * (len(result) // num_sheet)
            chunk = result.iloc[start:end]
            chunk.to_excel(writer, sheet_name='Sheet_{}'.format(i))

        writer.close()

    input('File exported successfully as "EFormatted.xlsx" in the same folder. Press any key to quit.')

else:
    input('Invalid number, please restart the program and try again.')