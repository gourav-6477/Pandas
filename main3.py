
import math
import pandas as pd
def insert_row_after(purana_df, insert_index, new_row):
    """
    Inserts a row after a given row in a Pandas DataFrame.

    Args:
        df (pd.DataFrame): The original DataFrame.
        insert_index (int): The index of the row after which to insert the new row.
        new_row (dict): A dictionary representing the new row,
                      where keys are column names and values are the row values.

    Returns:
        pd.DataFrame: A new DataFrame with the row inserted.
    """

    # Create a DataFrame for the new row
    new_row_df = pd.DataFrame([new_row])

    # Slice the DataFrame
    df_before = purana_df.iloc[:insert_index + 1]
    df_after = purana_df.iloc[insert_index + 1:]

    # Concatenate the slices with the new row
    df_inserted = pd.concat([df_before, new_row_df, df_after], ignore_index=True)

    return df_inserted
df = pd.read_excel('C:/Users/Gourav/Desktop/Project2/Book10.xlsx')
new_df = pd.read_excel('C:/Users/Gourav/Desktop/Project2/Splitdata.xlsx')
no_change_array = []
change_array = []
column_list = list(df.columns)
for j in range(51):
    if column_list[j][0] == "N":
        no_change_array.append(j)
    else:
        change_array.append(j)

new_column_list = list(new_df.columns)
pd.options.display.max_columns = None
i = 0
for number in change_array:
    column_name = new_column_list[number]
    new_df[column_name].fillna(0)
while i < len(new_df):
    closing_quantity_value = new_df.iloc[i,24]
    physical_quantity_value = new_df.iloc[i,35]
    ratio_number = physical_quantity_value/closing_quantity_value
    nayi_row_add_krunga = {}
    for l in range(50):
        if l in no_change_array:
            nayi_row_add_krunga[new_column_list[l]] = new_df.iloc[i,l]
        else:
            value = new_df.iloc[i, l]
            if type(value) == str:
                nayi_row_add_krunga[new_column_list[l]] = 0
            else:
                nayi_row_add_krunga[new_column_list[l]] = new_df.iloc[i,l] * ratio_number
    nayi_row_add_krunga[new_column_list[24]] = physical_quantity_value
    nayi_row_add_krunga[new_column_list[25]] = physical_quantity_value
    nayi_row_add_krunga[new_column_list[35]] = physical_quantity_value
    nayi_row_add_krunga[new_column_list[37]] = 0
    new_df = insert_row_after(new_df,i,nayi_row_add_krunga)
    i += 2
    print(i)
i = 0
while i < len(new_df):
    store_array = []
    proportion_number = []
    new_row = {}
    j = 51
    while type(new_df.iloc[i,j]) != str and not math.isnan(new_df.iloc[i,j]):
        if j%2 != 0:
            store_array.append(new_df.iloc[i,j])
        else:
            proportion_number.append(new_df.iloc[i,j])
        j += 1
        if j > len(new_column_list)-1:
            break

    total = sum(proportion_number)
    dictionary = {}
    for number in change_array:
        dictionary[number] = new_df.iloc[i+1,number]
    for t in range(len(store_array)):
        new_row["SN"] = f'{new_df.iloc[i+t+1,0]}A'
        for k in range(1,50):
            if k in no_change_array:
                new_row[new_column_list[k]] = new_df.iloc[i, k]
            else:
                value = new_df.iloc[i, k]
                if type(value) == str:
                    new_row[new_column_list[k]] = 0
                else:
                    value -= new_df.iloc[i+1,k]
                    if t != len(store_array)-1:
                        meri_value = value * (proportion_number[t] / total)
                        if math.isnan(meri_value):
                            pass
                        else:
                            meri_value = round(meri_value)
                            dictionary[k] += meri_value
                            new_row[new_column_list[k]] = meri_value
                    else:
                        new_row[new_column_list[k]] = new_df.iloc[i, k] - dictionary[k]

        new_row["Store"] = store_array[t]
        new_row["Transfer Quantity"] = proportion_number[t]
        new_df = insert_row_after(new_df, i + t + 1, new_row)
    i += len(store_array) + 2
    print(i)
new_df.to_excel('output7.xlsx')












