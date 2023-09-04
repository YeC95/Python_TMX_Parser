import xml.etree.ElementTree as ET #For manipulating XML
import re #Regex
import html #For dealing with escaped html tags
import pandas as pd #To utilize Dataframe from this lib for formatting table-like data structure

def clean_segment(segment):
    if segment is None:
        return ""
    # Remove XML elements
    cleaned_segment = re.sub(r'<.*?>', '', segment)
    # Remove placeholders like %...%
    cleaned_segment = re.sub(r'%[^%]*%', '', cleaned_segment)
    # Contain escaped html tags
    cleaned_segment = html.unescape(cleaned_segment)
    # Remove extra spaces and newlines
    cleaned_segment = ' '.join(cleaned_segment.split())

    return cleaned_segment


input_tmx_file = "D:\\PythonTMX\\input\\tmx-file.tmx"
output_spreadsheet_file = "D:\\PythonTMX\\output\\output_cleaned_tmx_file1.xlsx"


tree = ET.parse(input_tmx_file) #ET object creation
root = tree.getroot() # Retrieve root element

# To keep track of languages encountered in the XML and stores cleaned text segments in a structured manner
languages = set()
cleaned_data = []

# Finding corresponding <seg> element within each <tu>, and print out the text content of the <seg> element.
for tu in root.iter('tu'):
    cleaned_segment_data = {} # Dictionary for tu pairs
    for tuv in tu.iter('tuv'): # Iterating <tuv> within the <tu>
        lang = tuv.get('{http://www.w3.org/XML/1998/namespace}lang')  # Predefined XML namespace. Using the .get() method to extract the value of the 'lang' attribute (langauge code)
        if lang:
            languages.add(lang) # Checks if there is a language attribute (lang) associated with the current 'tuv' element
            segment_element = tuv.find('seg')
            if segment_element is not None: # Checks if there is a 'seg' element within the current 'tuv' element (tuv.find('seg'))
                segment = segment_element.text # .text property in an ET contains text content within the element's tags
                cleaned_segment = clean_segment(segment) # Calling clean_segment function
                cleaned_segment_data[lang] = cleaned_segment # Storing cleaned segments in a dictionary, with 'lang' as the key
    if cleaned_segment_data:
        cleaned_data.append(cleaned_segment_data) # Only append non-empty dictionaries

print("Languages:", languages)  # Debug, to verify language extraction
print("Cleaned Data:", cleaned_data)  # To verify cleaned data

languages = sorted(languages) # Sorts the unique language codes stored in the languages set
df_dict = {lang: [data.get(lang, '') for data in cleaned_data] for lang in languages} # creating DataFrame dictionary, key is language code
df = pd.DataFrame(df_dict) # Creating a Pandas Dataframe from df_dict. Dataframe organizes the language into columns

with pd.ExcelWriter(output_spreadsheet_file) as writer:
    df.to_excel(writer, index=False) #writes the Excel file

print("Spreadsheet file created:", output_spreadsheet_file)
