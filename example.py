import requests

import xlwt

api_key = "g9J7GzJxorn5bsidYp6O8jya6MOicyPbfmjNh003"
docketId = "HHS-OS-2022-0012"

def GetComment(object_id):
    cnt = 0
    comments_set = []
    url_commentlist = "https://api.regulations.gov/v4/comments?filter[commentOnId]=" \
                      + object_id + "&page[size]=250&api_key=" + api_key
    resp = requests.get(url_commentlist)
    resp_dict = resp.json()
    if resp_dict["meta"]["numberOfElements"] == 0:
        return comments_set
    page_num = resp_dict["meta"]["totalPages"]
    for cmt_page in range(page_num):
        url_commentlist_page = "https://api.regulations.gov/v4/comments?filter[commentOnId]=" \
                               + object_id + "&page[number]=%d" % (cmt_page+1) \
                               + "&page[size]=250&api_key=" + api_key
        resp_page = requests.get(url_commentlist_page)
        resp_page_dict = resp_page.json()
        for row in range(len(resp_page_dict["data"])):
            comment_id = resp_page_dict["data"][row]["id"]
            url_single_document = "https://api.regulations.gov/v4/comments/"+comment_id+"?api_key="+api_key
            resp_single_comment = requests.get(url_single_document)
            single_comment = resp_single_comment.json()
            comments_set.append(single_comment["data"])
            cnt += 1
            if cnt>=100:
                return comments_set
    return comments_set


def save_comments(comments_set, sheet):
    sheet.write(0,0,"CommentOnDocumentID")
    sheet.write(0, 1, "name")
    sheet.write(0, 2, "LastModifyDate")
    sheet.write(0, 3, "Comment")
    for idx in range(len(comments_set)):
        cmt = comments_set[idx]
        doc_id = cmt["attributes"]["commentOnDocumentId"]
        name = cmt["attributes"]["firstName"]+cmt["attributes"]["lastName"]
        content = cmt["attributes"]["comment"]
        date = cmt["attributes"]["modifyDate"]
        sheet.write(idx+1, 0, doc_id)
        sheet.write(idx + 1, 1, name)
        sheet.write(idx + 1, 2, content)
        sheet.write(idx + 1, 3, date)


if __name__ == "__main__":
    comments_set = []

    url_documentlist = "https://api.regulations.gov/v4/documents?filter[docketId]="+docketId+"&api_key="+api_key
    resp = requests.get(url_documentlist)
    response_dict = resp.json()
    for page in range(response_dict["meta"]["totalPages"]): ## get page# of the documents
        url_document_page = "https://api.regulations.gov/v4/documents?filter[docketId]="\
                              +docketId+"&page[number]="+"%d"%(page+1)+"&api_key="+api_key
        resp = requests.get(url_document_page)
        response_doc_dict = resp.json()
        for row in range(len(response_doc_dict["data"])): ## get single document in every page
            single_document = response_doc_dict["data"][row]
            object_id =  single_document["attributes"]["objectId"]
            comments_doc_set = GetComment(object_id)
            if len(comments_doc_set) == 0:  ## if no comment under the document, continue
                continue
            else:
                comments_set.extend(comments_doc_set)
    ## check document(0001)
    url_single_document = "https://api.regulations.gov/v4/documents/"\
                          +docketId+"-0001"+"?api_key="+api_key
    resp = requests.get(url_single_document)
    if resp.status_code != "404":
        response_doc_dict = resp.json()
        single_document = response_doc_dict["data"]
        document_id = single_document["id"]
        document_link = single_document["links"]["self"]
        doc_type = single_document["attributes"]["documentType"]
        mod_date = single_document["attributes"]["modifyDate"]
        object_id = single_document["attributes"]["objectId"]
        comments_doc_set = GetComment(object_id)
        if len(comments_doc_set) != 0:
            comments_set.extend(comments_doc_set)

    file = xlwt.Workbook("encoding = utf-8")
    sheet = file.add_sheet("comments")
    save_comments(comments_set, sheet)
    file.save("comments.xls")














