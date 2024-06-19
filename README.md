
# Automate Excel Reporting Using Python

There’s a lot of pain points in Excel that make it a tool that’s cumbersome and repetitive for data manipulation. But did you know that you can also use Python to automate repetitive tasks in Excel? In this video, I will be sharing my favorite ways to automate Microsoft Excel using Python. In particular, we will be using the open-source tools: Pandas, xlwings & plotly. After this video, you will be able to create a custom Python script that allows you to combine excel files & create charts out of them.


## Video Tutorial
[![YouTube Video](https://img.youtube.com/vi/JoonRjMsSdY/0.jpg)](https://youtu.be/JoonRjMsSdY)

## Changes after releasing the video
**Please note**<br/>
With pandas version 1.4.0 DataFrame.append() and Series.append() have been deprecated and will be removed in a future version.<br/>
Hence, I have changed the code as follows to merge all Excel files into one DataFrame:
```diff
- df = df.append(pd.read_excel(file), ignore_index=True)

+ df_tmp = pd.read_excel(file)
+ df = pd.concat([df, df_tmp], ignore_index=True)
```

## 🤓 Check Out My Excel Add-ins
I've developed some handy Excel add-ins that you might find useful:

- 📊 **[Dashboard Add-in](https://pythonandvba.com/grafly)**: Easily create interactive and visually appealing dashboards.
- 🤪 **[Emoji Add-in](https://pythonandvba.com/emojify)**: Add a touch of fun to your spreadsheets with emojis.
- 🛠️ **[MyToolBelt Add-in](https://pythonandvba.com/mytoolbelt)**: A versatile toolbelt for Excel, featuring:
  - Creation of Pandas DataFrames and Jupyter Notebooks from Excel ranges
  - ChatGPT integration for advanced data analysis
  - And much more!

## 🤝 Connect with Me
- 📺 **YouTube:** [CodingIsFun](https://youtube.com/c/CodingIsFun)
- 🌐 **Website:** [PythonAndVBA](https://pythonandvba.com)
- 💬 **Discord:** [Join the Community](https://pythonandvba.com/discord)
- 💼 **LinkedIn:** [Sven Bosau](https://www.linkedin.com/in/sven-bosau/)
- 📸 **Instagram:** [sven_bosau](https://www.instagram.com/sven_bosau/)

## Support 
If you appreciate the project and wish to encourage its continued development, consider [supporting my work](https://pythonandvba.com/coffee-donation).
[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://pythonandvba.com/coffee-donation)

## Feedback & Collaboration
For feedback, suggestions, or potential collaboration opportunities, reach out at contact@pythonandvba.com.
![Logo](https://www.pythonandvba.com/banner-img)
