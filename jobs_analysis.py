import pandas as pd
import numpy as np
import re
import pathlib

# Get current path
cur_path = str(pathlib.Path().absolute())

# Import dataset
df = pd.read_excel(cur_path + "\\jobs.xlsx")
df.columns = np.arange(0,df.shape[1])

# New dataframe to store transformations
df_clean = pd.DataFrame()

# Iterate over columns to transform data
for c in df.columns:
    
    # Drop na
    offer = df[df.columns[c]].dropna()
    offer = offer.reset_index(drop=True)
    
    # Get index where information of interest starts    
    info_start = np.sum(offer[offer.str.startswith("Company Name") == True].index)
    offer = offer.iloc[info_start:].reset_index(drop=True)
    
    # Separate company name and location    
    offer[0] = offer.iloc[0].replace("Company Name", "").replace(" Company Location ", "<>")
    
    # Drop unnecessary rows
    offer.drop(np.arange(1,11), inplace=True)
    offer.drop([12, 15, 16, 17], inplace=True)
    offer = offer.reset_index(drop=True)
    
    # Drop rows if there is info about someone that posted the offer
    offer = offer.drop(np.arange(4, 9)) if (np.sum(offer.str.contains("Send InMail")) != 0) else offer
    offer = offer.reset_index(drop=True)
    
    # Lowercase
    offer[4] = offer.iloc[4:].str.cat(sep="_").lower()
    
    # Replace unnecessary characters for spaces
    offer[4] = re.sub(r"[,.:;_()!?¿@#/&%*=…]", " ", offer[4])

    # Transpose series and create provisory dataframe
    offer_clean = pd.DataFrame(offer.iloc[0:5]).reset_index(drop=True).T
    
    # Split company name and company location into separate columns
    offer_clean[[5, 6]] = offer_clean[0].str.split("<>", expand=True)
    offer_clean.drop(0, axis=1, inplace=True)

    # Rename columns    
    offer_clean.columns = ["seniority", "size", "field", "description", "company", "location"]
    
    # Reorder columns
    offer_clean = offer_clean[["company", "location", "size", "field", "seniority", "description"]]
    
    # Append to transformed dataframe
    df_clean = df_clean.append(offer_clean, ignore_index=True)


# Drop duplicates - this is done after transformations because some contents 
# of the page can change even with the same job offer
df_clean.drop_duplicates(inplace=True)

## Text analysis

# Read glossary dataset
df_glossary = pd.read_excel(cur_path + "\\glossary.xlsx", names=["tool"])
df_glossary.drop_duplicates(inplace = True)

# Add spaces before and after each word - necessary because of terms like 'r', 
# 'sap' that might appear as characters of other words
df_glossary = " " + df_glossary["tool"] + " "

# Convert series to list
glossary = df_glossary.to_list()

# Blank dictionary
dict_containsword = {}

# Iterate over glossary list
for i in glossary:
    
    # For each term, check if it is present on a job offer at least once
    dict_containsword[i] = np.sum(df_clean["description"].str.contains(i))
    

# Generate dataframe with counting results
df_result = pd.DataFrame.from_dict(dict_containsword, orient="index", columns=["count"])

# Sort values by number of occurences of each term of the glossary
df_result.sort_values("count", ascending=False, inplace=True)

# Plot bars with most occurences
ax = (df_result*100/df_clean.shape[0]).head(30).\
      plot(kind="barh", 
           legend=None, 
           figsize=(16,16), 
           xlim=(0,100), 
           xticks=np.arange(0,101,10), 
           title="Percentage of job offers that include a term")
ax.invert_yaxis()
