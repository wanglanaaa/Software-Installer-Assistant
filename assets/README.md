# 软件安装助理图标

- `app-icon.png`：使用内置 imagegen 生成的透明背景原图。
- `app-icon.ico`：由原图转换的 Windows 图标，含 16、20、24、32、40、48、64、128、256 像素尺寸，用于程序文件及标题栏。界面 LOGO 直接嵌入 PNG 原图，并采用高质量缩放，避免放大小图标造成模糊。
- 格式转换：`./scripts/convert-app-icon.ps1`。图标嵌入程序集，运行时不依赖外部图片文件。

生成提示词（内置工具模式）：

> Use case: logo-brand. Create one polished Windows desktop application icon for a software installation manager named 软件安装助理. No text, no letters, no words. A bold simple white software package / open box silhouette with one integrated download arrow and a small clear checkmark, on a softly rounded square tile in deep blue with a restrained teal accent. Clean professional Windows utility aesthetic, near-flat with very subtle depth, not photorealistic, no scene, no mockup. Strong geometric silhouette legible at 16x16 and 32x32 pixels; minimal details, avoid thin lines or tiny decorations. Front-on centered composition, tile occupies about 88 percent of square canvas, equal small transparent margins on all four sides, actual transparent background outside the tile. Deliver a single square icon asset.
