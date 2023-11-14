# EmailBot
Designed an auto email sender

# Interface
<img width="400" alt="image" src="https://github.com/YujunZeng-Lulu/EmailBot/assets/57317964/df91d066-6938-47d6-8bf7-aca2ecc2a908">

# Instructions
0. Strongly suggest sending the email to yourself before sending it to others for a double check!!!
1. Email servers: QQ - smtp.qq.com 456, QQ Enterprise - smtp.exmail.qq.com 465, 163 - smtp.163.com 25.
2. Server, From(sender address), Password and signature info are required, while others are optional.
3. Two modes: 
 i. Send to one person: fill in To(receiver's address), Subject and Message, and leave Filepath empty.
 ii. Send to a group: select an Excel file, which contains the receivers' name/address/firm name/etc (see test.xls for reference).
4. For mode ii, the only required column in the Excel is receivers' email address. The column name must be one of the following: <recipient>/<email>/<to>.
5. For mode ii, all other columns will be matched and filled to both subject and message. The column name must be in the form <XXX>, and the same form in the subject/message.（进一步解释：群发模式下，邮件模版是可以灵活定制的。只要在邮件模版里用<>进行挖空，然后在Excel里把<>相关信息补上即可。）
6. The final subject would be the text in 'or Type Your Subject:' (with <xxx> filled by info in the Excel, if applicable).
7. There are some fields (prompt, generate) retained for future GPT usage, which are not yet available given concerns issues regarding ChatGPT API（1. 多人使用的情况下，最好能用一个公用的GPT账号，防止私人账号被封号；2. OpenAI需要收费，约每封邮件0.1美元。 ）.
8. To add new templates, please find the file templates.json and put them there.
9. To add new subject templates, please find subjects.json.
10. To change any format of the signature, please find signature.html.
11. Met a bug or need help? Please contact Yujun (yujun.lulu.zeng@gmail.com).

Wish you a bright day! :)

# Reference
Part of the code was suggested by ChatGPT
