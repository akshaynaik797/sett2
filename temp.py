record = [('606', 2, 'religare', 'Max PPT', 'Payment against Claim Reference Number:91389882-02 Policy No :\r\n 11315557  Proposer Name :BHARAT BHUSHAN GUPTA Patient Name :BHARAT BHUSHAN\r\n GUPTA'), ('617', 1, 'Medi_Assist', 'Max PPT', 'Settlement of your Claim Reference :CCN# 101319017 under policy#\r\n 12120034200400000013'), ('626', 1, 'fhpl', 'Max PPT', 'Cashless Settlement Letter : Patient Name : Suman    (Wife) ; Employee Name : Arvind  Kumar  ;  Employee ID : 1016'), ('659', 1, 'Raksha', 'Max PPT', 'Claim Settlement Letter From Raksha Health Insurance TPA Pvt.Ltd.\r\n (UIC54517833815,9556146,SHAILANDRA KUMAR MISHRA.)'), ('752', 1, 'Raksha', 'Max PPT', 'Claim Settlement Letter From Raksha Health Insurance TPA Pvt.Ltd.\r\n (M58ADD676ILBS,9331198,Dr. Aditi Goel.)')]
for i in record:
    with open("records.csv", "a+") as fp:
        i = str(i).replace("(", "").replace(")", "")
        fp.write(i + '\n')