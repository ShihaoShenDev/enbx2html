# enbx2html
将希沃白板课件转换为 HTML
![image](https://github.com/user-attachments/assets/54b19dea-a048-4a5d-ac5b-3e888a85e458)
![image](https://github.com/user-attachments/assets/9f0a4e66-6506-4565-bd67-e0141016beca)

## 使用方法
```bash
enbx2html.py [-h] [-o OUTPUT] [--info] input_file
```

## 示例
```bash
enbx2html.py example.enbx
Unzipping D:\source\EasiNote\example.enbx...
Converting from D:\source\EasiNote\example_html to D:\source\EasiNote\example_html...
Document Metadata: {'Name': '新手引导-排版(05-30 0039 修改)', 'Creator': 'hxzryxxsujtntpyxwosyuvx56p81p7s1', 'CreatedDateTime': '05/26/2022 06:31:18', 'ModifiedDateTime': '05/30/2022 06:21:59'}
Board parsed: {'width': 1280.0, 'height': 720.0}, 8 slides found.
References parsed: 64 resources found.
Mapped 8 slide files.
Source and Output are the same directory. Skipping resource copy.
HTML generation complete.
Done! Output at: D:\source\EasiNote\example_html\index.html
```

## 附注

### 原理
众所周知，`enbx` 是 [OPC](https://baike.baidu.com/item/%E5%BC%80%E6%94%BE%E6%89%93%E5%8C%85%E7%BA%A6%E5%AE%9A/22723424) 格式（怎么知道的？[见这里](#thanks)）

于是我就告诉 AI 这个思路，让 AI 来写。
![image](https://github.com/user-attachments/assets/cb8e2753-8377-4ccb-8160-c7f9b6afc9c4)

~~结论：反编译了个寂寞~~

### Thanks
[@lindexi](https://github.com/lindexi)
![image](https://github.com/user-attachments/assets/2ff03add-19cb-422f-a984-6f9ba9456374)

本项目在 [Trae](https://trae.ai) 的帮助下开发。
