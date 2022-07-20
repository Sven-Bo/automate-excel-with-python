
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

## Author
- Sven from Coding Is Fun
- YouTube: https://youtube.com/c/CodingIsFun
- Website: https://pythonandvba.com


## Feedback
If you have any feedback, please reach out to me at contact@pythonandvba.com


![Logo](https://www.pythonandvba.com/banner-img)

