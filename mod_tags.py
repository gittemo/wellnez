import pandas as pd
import string

def excel_cols(num_cols):
    col_names = []
    letters = string.ascii_uppercase

    for i in range(num_cols):
        if i < 26:
            col_names.append(letters[i])
        else:
            col_names.append(letters[i//26 - 1] + letters[i%26])
    return col_names

def transform_excel(input_filepath, output_filepath):
    df = pd.read_excel(input_filepath)
    df.columns = excel_cols(len(df.columns))
    df = df.drop(df.index[:4])

    shop = pd.DataFrame()

    shop['Title'] = df['E']
    shop['Body (HTML)'] = df['EV']
    shop['Vendor'] = df['D']
    shop['Product Category'] = 'Home & Garden > Pool & Spa > Saunas'
    shop['Type'] = df['FB']
    shop['Published'] = 'TRUE'
    shop['Variant SKU'] = df['C']
    shop['Variant Price'] = df['FK']
    shop['Variant Compare At Price'] = '9.99'
    shop['Variant Requires Shipping'] = 'TRUE'
    shop['Gift Card'] = 'FALSE'
    shop['SEO Title'] = df['E']
    shop['SEO Description'] = df['EW']
    shop['Google Shopping / Google Product Category'] = 'Home & Garden > Pool & Spa > Saunas'
    shop['Status'] = 'active'
    shop['Tags'] = df['EX'].str.replace(' ', ', ')

    shop['Image Src'] = df['FF'].str.split(',').fillna('').apply(lambda x: [url for url in x if url])
    shop = shop.explode('Image Src')
    shop['Image Position'] = shop.groupby('Variant SKU').cumcount() + 1
    shop = shop.reset_index(drop=True)

    # Make a copy of the dataframe for processing
    shop_processed = shop.copy()

    # Iterate over unique SKUs (each product should have a unique SKU)
    for sku in shop['Variant SKU'].unique():
        # Get a boolean mask where the SKU matches and the image position is not 1
        mask = (shop_processed['Variant SKU'] == sku) & (shop_processed['Image Position'] != 1)

        # For matching rows, retain only 'Title', 'Image Src' and 'Image Position'
        shop_processed.loc[mask, :] = shop_processed.loc[mask, ['Title', 'Image Src', 'Image Position']]

    # Save the processed dataframe to a CSV file
    shop_processed.to_csv(output_filepath, index=False)

if __name__ == "__main__":
    input_filepath = 'input.xlsx'  # replace with your input file path
    output_filepath = 'output2.csv'  # replace with your output file path
    transform_excel(input_filepath, output_filepath)
