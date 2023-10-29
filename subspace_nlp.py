!python -m spacy download en_core_web_lg
import pandas as pd
import spacy
from spacy.matcher import Matcher
nlp = spacy.load("en_core_web_lg")
nlp_2 = spacy.load("en_core_web_lg")
nlp_3 = spacy.load("en_core_web_lg")
nlp_4 = spacy.load("en_core_web_lg")
nlp_5 = spacy.load("en_core_web_lg")
nlp_6 = spacy.load("en_core_web_lg")
matcher = Matcher(nlp.vocab)
matcher_no = Matcher(nlp.vocab)
matcher_2 = Matcher(nlp_2.vocab)
matcher_3 = Matcher(nlp_3.vocab)
matcher_address = Matcher(nlp_4.vocab)
matcher_company = Matcher(nlp_5.vocab)
matcher_jobdesc = Matcher(nlp_6.vocab)

data = pd.read_excel(r'MyContacts.xlsx', engine='openpyxl', dtype = {'parsedTxt': object, 'fullname': str, 'company': str, 'job_title': str, 'address': object, 'phone': object, 'phone_2': object, 'email': object, 'email_2': object, 'website': object})
columns = ['fullname', 'company', 'job_title', 'address', 'phone', 'phone_2', 'email', 'email_2', 'website']
actual_data = pd.DataFrame(columns=columns)
n=0
data_list = []
eliminator = []

for i in range(171):
    data_to_append = {'fullname': None, 'company': None, 'job_title': None, 'address': None, 'phone': None, 'phone_2': None, 'email': None, 'email_2': None, 'website': None}
    data_to_append = {}
    splitted_data = str(data.iloc[i,0]).split("\n")

    one_sentence_data = str(data.iloc[i,0])
    doc = nlp_2(one_sentence_data)

    for j in splitted_data:
        if ("www" in j or "WWW" in j):
            pattern_url = [{"LIKE_URL": True}]
            matcher_2.add("URL", [pattern_url])
            matches_2 = matcher_2(doc)
            #print(matches_2)
            if matches_2 != [] and ("www" in str(doc[matches_2[0][1]:matches_2[0][2]]) or "WWW" in str(doc[matches_2[0][1]:matches_2[0][2]])):
                data_to_append['website'] = doc[matches_2[0][1]:matches_2[0][2]]

    one_sentence_data = str(data.iloc[i,0])
    one_sentence_data_removed = one_sentence_data.replace('\n', ' ')
    doc = nlp(one_sentence_data)
    doc_name = nlp_3(one_sentence_data_removed)
    doc_no = nlp_2(one_sentence_data_removed)
    doc_address = nlp_4(one_sentence_data_removed)
    doc_company = nlp_5(one_sentence_data_removed)
    doc_jobdesc = nlp_6(one_sentence_data_removed)
    pattern_mail = [{"LIKE_EMAIL": True}]
    pattern_no = [{"ORTH": "+91", "OP": "*"}, {"ORTH": " ", "OP": "*"}, {"ORTH": "-", "OP": "*"}, {"LIKE_NUM": True}, {"ORTH": " ", "OP": "*"}, {"LIKE_NUM": True}]
    pattern_name = [{"ORTH": "Mr. ", "OP": "*"}, {"ORTH": "Mrs. ", "OP": "*"}, {"POS": "PROPN", "ENT_TYPE": "PERSON", "OP": "+"}]
    pattern_address = [{"POS": "PROPN", "ENT_TYPE": "GPE", "OP": "*"}, {"ORTH": ", ", "OP": "*"}, {"POS": "PROPN", "ENT_TYPE": "GPE", "OP": "*"}]
    pattern_jobdesc = [{"POS": "NOUN"}]
    pattern_company = [{"POS":"PROPN", "OP": "+"}]
    matcher.add("EMAIL_ADDRESS", [pattern_mail])
    matcher_3.add("FULL_NAME",[pattern_name], greedy='LONGEST')
    matcher_address.add("ADDRESS", [pattern_address], greedy='LONGEST')
    matcher_company.add("COMPANY", [pattern_company], greedy = 'LONGEST')
    matcher_jobdesc.add("JOB TITLE", [pattern_jobdesc], greedy = 'LONGEST')
    matcher_no.add("PHONE_NO", [pattern_no])
    matches = matcher(doc)
    matches_no = matcher_no(doc_no)
    matches_3 = matcher_3(doc_name)
    matches_address = matcher_address(doc_address)
    matches_jobdesc = matcher_jobdesc(doc_jobdesc)
    matches_company = matcher_company(doc_company)

    for k in range(len(matches)):
        if matches[k] != []:
            if "@" in str(doc[matches[k][1]:matches[k][2]]):
                if k > 0:
                    data_to_append['email_2'] = doc[matches[k][1]:matches[k][2]]
                else:
                    data_to_append['email'] = doc[matches[k][1]:matches[k][2]]


    for k in range(len(matches_3)):
        if matches_3[k] != []:
            span = doc_name[matches_3[k][1]:matches_3[k][2]]
            exclusion_list = ["-",".","@","1","th","Road","Towers","Tower","Nadu","Pvt","Limited","Ltd","Plaza","Complex","Dist","Workspace"]
            if (not any(char.isdigit() for char in str(span))) and (not any(char in exclusion_list for char in str(span))) :
                #print(span)
                count_1 = count_2 = 0
                for token in span:
                    count_1 += 1
                    if (token.ent_type_ == "PERSON" and token.text not in exclusion_list) and (not any(char in exclusion_list for char in token.text)) or (token.text in ["Chauhan","Jain"]):
                        count_2 += 1
                if count_1 == count_2:
                    data_to_append['fullname'] = str(span)


    for k in range(len(matches_address)):
        if matches_address[k] != []:
            span = doc_address[matches_address[k][1]:matches_address[k][2]]
            data_to_append['address'] = str(span.text)

    for k in range(len(matches_jobdesc)):
        if matches_jobdesc[k] != []:
            span = doc_jobdesc[matches_jobdesc[k][1]:matches_jobdesc[k][2]]

    for k in range(len(matches_company)):
        if matches_company[k] != []:
            span = doc_company[matches_company[k][1]:matches_company[k][2]]
            dontadd = []
            for token in span:
                inclusion_list = ["Pvt","Limited","Ltd",]
                if token.text in inclusion_list:
                    dontadd.append(span)
                    for m in range(len(dontadd)):
                        index_1 = str(span.text).find("Pvt")
                        index_2 = str(span.text).find("Limited")
                        index_3 = str(span.text).find("Ltd")
                        index_4 = max(index_1, index_2)
                        index = max(index_4, index_3)
                        eliminator.append(str(span.text)[:(index-1)])
                        find = 0
                        for eliminate in range(len(eliminator)-1):
                            if str(span.text)[:(index-1)] == str(eliminator[eliminate]):
                                find += 1
                        if find == 0:
                            data_to_append['company'] = str(span.text)[:(index-1)]
                            n += 1

    data_list.append(data_to_append)
    got_it = 0
    more = 0
    for k in range(len(matches_no)):
        if matches_no[k] != []:
            span = doc_no[matches_no[k][1]:matches_no[k][2]]
            if len(str(span.text)) > 10:
                more += 1
                if more == 1:
                    data_to_append['phone'] = str(span.text)
                if more == 2:
                    data_to_append['phone_2'] = str(span.text)
           

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

actual_data = pd.DataFrame(data_list)
print("THE FINAL TABLE AFTER EXTRACTION IS:")
print(actual_data)

pd.reset_option('display.max_rows')
pd.reset_option('display.max_columns')
pd.reset_option('display.width')
