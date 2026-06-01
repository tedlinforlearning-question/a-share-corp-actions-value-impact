# Excel 模板

模板文件已上传到 GitHub，请直接下载：

| 模板 | 直接下载 |
|:---|:---|
| 通用模板 | https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/DCF_Model_Template.xlsx |
| IPO 专用 | https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/DCF_Model_Template_IPO.xlsx |
| 并购买方 | https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/DCF_Model_Template_MA_Buy.xlsx |
| 并购卖方 | https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/DCF_Model_Template_MA_Sale.xlsx |

右键链接 → "另存为" 即可下载。

---

## 本地生成

也可以使用 Python 脚本本地生成：

```bash
# 生成全部三种场景模板
python ../scripts/build_dcf_model.py --all

# 仅生成 IPO 模板
python ../scripts/build_dcf_model.py --scenario ipo
```
