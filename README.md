
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

## Learn Excel Automation with Python
If this repo helped you, my [Excel Automation Course](https://pythonandvba.com/excel-automation-course/) teaches the full workflow from zero: Python for Excel users, xlwings, pandas and real projects.

Also check out my other [tools and templates](https://pythonandvba.com/solutions).

## Connect with Me
- **YouTube:** [CodingIsFun](https://youtube.com/c/CodingIsFun)
- **Website:** [PythonAndVBA](https://pythonandvba.com)
- **LinkedIn:** [Sven Bosau](https://www.linkedin.com/in/sven-bosau/)
- **Contact:** [Get in Touch](https://pythonandvba.com/contact)
## Support
If you find this project helpful, consider buying me a coffee. 

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://pythonandvba.com/coffee-donation)
