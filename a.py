[32mInitiate Downloading....[0m
Reading Emails From EmailId.txt
Error accessing https://apim.people.deloitte/personresumes?email=prasarangi@deloitte.com: HTTPSConnectionPool(host='apim.people.deloitte', port=443): Max retries exceeded with url: /personresumes?email=prasarangi@deloitte.com (Caused by SSLError(SSLCertVerificationError(1, '[SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed: self-signed certificate in certificate chain (_ssl.c:1000)')))
Error accessing https://apim.people.deloitte/personresumes?email=spujari@deloitte.com: HTTPSConnectionPool(host='apim.people.deloitte', port=443): Max retries exceeded with url: /personresumes?email=spujari@deloitte.com (Caused by SSLError(SSLCertVerificationError(1, '[SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed: self-signed certificate in certificate chain (_ssl.c:1000)')))
Error accessing https://apim.people.deloitte/personresumes?email=rishabhsharma96@deloitte.com: HTTPSConnectionPool(host='apim.people.deloitte', port=443): Max retries exceeded with url: /personresumes?email=rishabhsharma96@deloitte.com (Caused by SSLError(SSLCertVerificationError(1, '[SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed: self-signed certificate in certificate chain (_ssl.c:1000)')))
Exception in Tkinter callback
Traceback (most recent call last):
  File "tkinter\__init__.py", line 1967, in __call__
  File "DeloitteResume_Downloader_Search_GUI.py", line 207, in Download_Resume
  File "DeloitteResume_Downloader_Search_GUI.py", line 68, in read_urls_from_file1
FileNotFoundError: [Errno 2] No such file or directory: 'ResumeDownloadURL.txt'
