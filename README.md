# SimpleTranslator(English translation tool)

[日本語 README はこちら][f]

![d](img/jobs02.gif)

I created a tool that can easily translate foreign words pasted into Excel using Selenium Basic and DeepL.

Setting up the environment is not easy, but once you are done setting up the environment, other than updating the Chrome Driver, there is no need to set up.

## ■ Background

I am currently working in a patent-related field and am using a search system to search for patent documents.

The patent documents that are output as search results are not only in Japanese but also in foreign languages such as English, and we sometimes read foreign language patent documents to understand the technical contents.

## ■ Theme

However, the translation accuracy of the translation tools provided is low, and it takes time to understand the translation results of foreign language patent documents.

## ■ Purpose

The idea is to provide a translation tool that can compare the translated text with the original text with high accuracy by displaying the original text in a foreign language and the translated text translated by DeepL side by side.

## ■ What you need

### OS

- Microsoft Windows 10

### Software

- Google Chrome
- Microsoft Excel 2016 or Microsoft Excel 2019
- Selenium Basic
- Chrome Driver

## ■ Setting up the environment to run Selenium Basic

If you have Selenium Basic installed and running, you do not need to do the following. Please proceed to the next step.

1. [Open the tutorial page on how to install SeleniumBasic and perform web scraping from Excel (VBA)][a].
2. Download and install Selenium Basic referring to [the above article][a].
3. Download the Chrome Driver by referring to [the article above][a], and then copy and overwrite the downloaded Chrome Driver to the folder where you installed Selenium Basic.
4. Check the operation of Selenium Basic referring to [the above article][a].
5. If the error occurs because ".Net Fremework 3.5" is not installed, refer to the link at [the bottom of the above article][a] to install.
6. If you are unable to install ".Net Fremework 3.5" at this point, perform the following step 7. 7.[How to install .NET3.5 on Windows 10! Install .Net Fremework 3.5 by referring to the page How to Install][b] . Net Fremework 3.5. Make sure to revert the changed registry values.

## ■ How to use Simple Translator

1. [1.SimpleTranslator zip file][e] ← Click to download.

2. Unzip the zip file downloaded in step 1 above, and open SimpleTranslator.xlsm in the SimpleTranslator folder.

3. Select the language selection combo box.

   ![d](img/en_normal_img001.png)

4. Select the target language.

   ![d](img/en_normal_img002.png)

5. Select the radio button to choose the display order of "Translated Text → Source Text" or "Source Text → Translated Text".

   ![d](img/en_normal_img003.png)

### How to use it in Web articles (For how to use it in patent documents and PDF, please skip this section and go to the following sections.)

6. Trace the English text displayed on [the sample page][c] with your mouse to copy it.

   ![d](img/en_normal_img004.png)

7. Right-click on sheet A2 in Excel, and select the second icon from the left under Paste Options.

   ![d](img/en_normal_img005.png)

8. Press the "Shaping button" to fill in the blank lines.

   ![d](img/en_normal_img006.png)

9. Press the "Translate (HTML)" button.

   ![d](img/en_normal_img007.png)

10. The "Save As" dialog box will appear.

    ![d](img/en_normal_img008.png)

11. Enter a name and save the file. In this case, enter "sample" and click the "Save button".

    ![d](img/en_normal_img009.png)

12. The translation will start and a progress bar will appear.

    ![d](img/en_normal_img010.png)

13. When the translation is complete, the browser will open and output the translation results.

    ![d](img/en_normal_img011.png)

14. The translation results will be created in HTML format in a directory in the same location as the application.

## ■ Caution

If Google Chrome is updated, Chrome Driver will no longer work.

In this case, please refer to [the above article][a] and overwrite the Chrome Driver that matches the version of Google Chrome.

## ■ Download page

1. [Selenium Basic](https://florentbr.github.io/SeleniumBasic/)
2. [Chrome Driver](https://chromedriver.chromium.org/downloads)

## ■ Reference page

1. [tutorial on how to install SeleniumBasic and perform web scraping from Excel (VBA)][a]
2. [How to install .NET3.5 on Windows 10!][b]
3. [sample][c]

[a]: https://lil.la/archives/3436
[b]: https://bgt-48.blogspot.com/2019/04/windows10net35.html
[c]: https://www3.nhk.or.jp/news/html/20210728/k10013161181000.html
[e]: https://github.com/masatofujiki/SimpleTranslator/archive/refs/tags/v1.1.0.zip
[f]: https://github.com/masatofujiki/SimpleTranslator/blob/main/README_JA.md
