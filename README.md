Useful VBAs is an add-in made to foot leadsheets in my work as an auditor. A leadsheet is a summary of financial information that acts as the starting point of an audit working paper. When originally imported from Caseware, the leadsheet is filled with hardcoded numbers which a person then has to manually put in formulas. When a leadsheet gets larger and larger, the manual process becomes more arduous. In order to make the process more efficient, I made an excel add-in which is being used by several members in my accounting firm. The challenge was in putting safeguards in to ensure that formulas do not overwrite incorrect cells and it only foots where appropriate. While this is an excel demo file that is an xlsm, the actually add-in is an xlam and uses the firms actual footing typography (for the purpose of the demo I just say 'footed').  I also added a grouping functionality to group together by sum and a list formula for Caseware IDEA extraction (Data analysis extraction). This was my second time using VBA for a project so there is definitely room to refactor and make the logic more succinct.

Before using Leadsheet:

![image](https://github.com/Tejasvshetty/Useful-VBAs/assets/78327281/931900f2-bd67-45e4-9706-808cbaabb46c)

After using Leadsheet:

![image](https://github.com/Tejasvshetty/Useful-VBAs/assets/78327281/bb435695-570e-4620-803a-885c6d555b25)

![image](https://github.com/Tejasvshetty/Useful-VBAs/assets/78327281/26a3c45e-4a50-41c6-a80d-fa02a80672a8)

Before Grouping:

![image](https://github.com/Tejasvshetty/Useful-VBAs/assets/78327281/384be57c-4fc9-41ed-985c-96ff0fa03951)

After Grouping:

![image](https://github.com/Tejasvshetty/Useful-VBAs/assets/78327281/fad5db81-8801-415b-ba62-a6da6487ee7a)




