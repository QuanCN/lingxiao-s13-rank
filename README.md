# 凌霄 S13 属性排名

微信可打开：https://quancn.github.io/lingxiao-s13-rank/

## 只换表格就能更新（推荐）

表格文件名必须是本仓库里的 **`data.xlsx`**（列顺序与原来一致：序号、成员名称、爵位、阵容、宫四兵、集结加成、步兵生命、步兵防御、弓兵攻击、弓兵破坏）。

### 方式一：电脑上一键更新

1. 从金山文档导出最新 Excel  
2. 覆盖本文件夹里的 `data.xlsx`（不要改名）  
3. 双击 **`更新排名.bat`**  
4. 等它推送完成，约 1 分钟后刷新网页链接  

首次若提示缺库，在本文件夹打开命令行执行：

```bat
pip install -r requirements.txt
```

### 方式二：网页上传表格（不装 Python）

1. 打开 https://github.com/QuanCN/lingxiao-s13-rank  
2. 点开 `data.xlsx` → 右上角上传/替换文件  
3. 提交后 GitHub Actions 会自动重算 `index.html` 并发布页面  

## 评分规则

属性分 = 0.55×步兵归一 + 0.45×弓兵归一（**不含**爵位、阵容、集结）