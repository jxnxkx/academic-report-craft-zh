# 图中文字重排

报告里的图是程序生成的位图，字体是思源黑体、DejaVu Sans 一类，要改成与正文一致的宋体加 Times New Roman。2026-11-02 用此方法处理了 5 张组合图，数据和版面位置都没有变。

## 先盘点

1. 解压 `.docx`，列出 `word/media/` 中的图片及像素尺寸。
2. 从 `document.xml` 读出每张图的 `wp:extent`（显示尺寸，EMU，914400 = 1 英寸）和 `a:srcRect`（裁剪）。被裁剪的图只显示一部分，要按显示的部分判断是否需要改；图题和实际显示的内容不符时告诉用户。
3. 校徽、照片、仿真截图、从文献截取的示意图不重画。

## 重画方式

整张图按原图像素坐标布局，换算成 matplotlib 的图幅比例；画布大小设为 Word 中的显示尺寸，`dpi=330` 输出，替换后 Word 中的位置和大小都不变。

| 面板类型 | 做法 |
| --- | --- |
| 照片、截图、示意图 | 从原图裁下，`imshow` 放回原位置 |
| 柱状图 | 按图上标注的数值重画，柱宽、坐标范围按原图测量 |
| 曲线 | 按颜色掩膜逐列取像素中位数，换算为数据坐标，轻度平滑后重画 |
| 散点、多条重叠的环线等不便取点的 | 只裁下坐标框内部的图形，`imshow` 到与原坐标一致的 `extent`，由 matplotlib 重画坐标轴、刻度和标签 |
| 照片上的步骤标签 | 盖一个圆角色块，写宋体白字，加 0.5 磅白色描边代替粗体 |

### 测量坐标

- 坐标轴：在面板范围内找暗色像素最多的列和行，即为左轴和下轴。用连续暗色段的端点确定坐标框。
- 刻度：在轴线外侧 4 至 6 像素处，找出连续非白色的短线段，取中心位置。用两个刻度的像素位置和数值算出比例尺，再推出坐标范围。
- 刻度标签有颜色（如双纵轴图）时，阈值用“非白色”，不要用“暗色”。

### 坐标框内的文字

图内注释（如“*E*<sub>max</sub> ≈ 2.10 mmHg/ml”“原示范起点”）要先从位图中擦掉：在文字框范围内把偏红的像素（`R − G > 4`）或整块空白区域改成白色；文字盖住的线条用上方一行像素补齐。然后用 matplotlib 在空白处重写。原位置压线时可以挪到空白处。

## 字体设置

```python
from matplotlib import font_manager as fm
for f in ['simsun.ttc', 'times.ttf', 'timesi.ttf']:
    fm.fontManager.addfont('C:/Windows/Fonts/' + f)
plt.rcParams.update({
    'font.family': ['Times New Roman', 'SimSun'],   # 西文 Times，中文回退宋体
    'font.size': 9,                                  # 小五，按插入后的尺寸
    'mathtext.fontset': 'custom', 'mathtext.rm': 'Times New Roman',
    'mathtext.it': 'Times New Roman:italic',
    'axes.unicode_minus': True,
})
```

- matplotlib 3.6 以后支持按字体列表回退，同一字符串里的中文和西文会自动分别使用宋体和 Times。
- **mathtext 不会回退到宋体。**“改变 $R_{\mathrm{S}}$ 得到的……”“$x$（m）”这类中文和数学混排的字符串，中文会显示成方框。要拆成几段分别排版，量出每段宽度后依次摆放；旋转 90° 的坐标轴标题同样处理，用 `rotation_mode='anchor'` 和 `va='baseline'`。画布 `dpi` 要和保存时一致，测量才准。
```python
# 分段排版：parts = [('(b) 改变 ', {}), (r'$R_{\mathrm{S}}$', {}), (' 得到的压力—容积环', {})]
def rich(fig, cx, cy, parts, rotation=0, fs_pt=9, dpi=330):
    from matplotlib.transforms import IdentityTransform
    r = fig.canvas.get_renderer(); fs = fs_pt * dpi / 72
    ts, adv = [], []
    for s, kw in parts:
        t = fig.text(0, 0, s, transform=IdentityTransform(), rotation=rotation,
                     rotation_mode='anchor', ha='left', va='baseline', **kw)
        e = t.get_window_extent(r); ts.append(t); adv.append(e.height if rotation else e.width)
    pos = -sum(adv) / 2
    for t, a in zip(ts, adv):
        t.set_position((cx + 0.36 * fs, cy + pos) if rotation else (cx + pos, cy - 0.36 * fs))
        pos += a
```

坐标轴标题先用同样长度的占位文字 `set_xlabel`，`fig.canvas.draw()` 后取它的位置，再隐藏占位、在该位置调用 `rich`。

- 变量写成斜体：`$x$`、`$E_{\mathrm{max}}$`；单位、数字用正体。
- 图例放不下时移到坐标框上方（`loc='lower center', bbox_to_anchor=(0.5, 1.0), ncol=2`），或减小 `handlelength`、`columnspacing`。横轴标签挤在一起时改成两行，与同图其他子图的写法保持一致。

## 放回文档

- 用 Python 的 `zipfile` 复制原 `.docx`，只把 `word/media/imageN.png` 换成新图，其余条目按原顺序写回。
- 导出 PDF，对比替换前后同一页的版面；页数应不变。
- 交付时另存为新文件；单张图也存一份到“图片_修改后”文件夹，方便对方自己插入。
