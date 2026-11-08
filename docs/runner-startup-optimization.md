# 软件集合 EXE 启动优化与验证

2026-11-08。只调整生成模板、资源封装和执行队列；管理工具的界面与操作流程未修改。开发阶段仅验证独立样例；用户确认后按 v1.2.4 更新模板并构建独立版本发布目录。

## 修改前的机制与瓶颈

1. `scripts/build-runner.ps1` 生成 .NET 9 / Windows x64 自包含单文件执行器，配置 `IncludeNativeLibrariesForSelfExtract=true`。模板约 109.44 MiB；用户安装资源还未计入该体积。此模式包含运行时，不能因为电脑装有 Framework 4.8 就自动复用它。
2. `BundleBuilder` 将按 SHA-256 去重的资源写入 ZIP（原为 `Fastest`），追加到执行器后，再写入长度、历史摘要字段和 magic 共 56 字节。生成时会读取、计算摘要和复制源文件，这发生在制作电脑，不是目标电脑启动阶段。
3. 普通启动先创建界面，然后后台定位尾部、读取 ZIP 目录和最多 4 MiB 的清单。原代码已跳过整包摘要，不做已安装软件/版本检查，也没有启动时提取全部安装包。不过它仍逐个验证所有资源路径、摘要格式和 ZIP 条目长度。
4. 原代码已按依赖顺序逐款准备和安装，只在安装阶段读取配置中明确指定的检测规则。无全盘软件扫描。此次保留这部分行为。
5. 原提取只比较文件长度，没有检查内容摘要；同长度损坏可进入执行阶段。原停止菜单默认等当前软件结束，退出可保留运行中的安装器；MSI 客户端没有纳入进程组。

本机样例中，旧版首次启动约 5.1–5.3 秒，而进入托管入口后到界面就绪约 0.4–0.46 秒。因此主要延迟位于程序装载/运行时启动这一段，不能归因于安装包全量解压。安全软件扫描可能贡献首次延迟，但没有关闭防护或做系统级跟踪，尚不能量化；专用运行库缓存目录未观测到释放文件，也不把全部延迟归因于运行库解压。

## 实施结果

- 正式模板构建路径改为 `net48`、AnyCPU；生成的 EXE 使用系统安装的 .NET Framework 4.8/4.8.1。Windows 7 SP1 可使用 4.8；.NET 9 仍只用于制作工具及现有测试引用。CLR 自身不嵌入、不释放。
- 复用同一份安装界面和运行逻辑，按条件链接必要源码，不移植制作工具。JSON 依赖 Newtonsoft.Json 13.0.3 的 Framework 程序集和 MIT 许可证内嵌在 EXE 中，直接从内存加载，不生成旁置 DLL。
- 开发测量模板标识为 `1.2.3+win7-demand`，v1.2.4 发布模板为 `1.2.4+win7-demand`。生成器拒绝同版本号的旧自包含模板，避免源码更新后仍悄悄使用旧模板。
- 启动只检查界面/选择所需的清单结构、安装参数、软件标识和依赖。资源路径、摘要格式、存在性和长度移到当前软件的提取阶段；未选软件不会提取或进行内容校验。
- 安装资源采用 ZIP 存储模式（不二次压缩），清单仍可压缩。保留内容去重；最终体积可能比快速压缩版本大，但提取不再消耗解压 CPU。
- 当前软件资源以 128 KiB 缓冲顺序复制，同时计算 SHA-256；限制写出长度，验证通过才启动。没有解压后的第二遍扫描。
- 合集的“取消安装”及安装中关闭窗口都取消准备、停止后续队列，并终止当前 Job Object 内的进程；等待受控子进程退出后才清理资源/关闭窗口。安装器先挂起，成功加入进程组并设置退出保护后才恢复；无法管理就阻止启动。共享引擎通过 `cancelCurrent: true` 启用此行为，制作工具原有调用不变。
- 超时也终止合集受控进程并暂停队列。MSI 客户端与受控子进程可取消，但 Windows Installer 服务可能继续回滚，不能保证系统服务任务已结束；此时报告待检查并保留资源，不结束系统中其他安装任务。

## 运行条件

新合集需要目标机**预装 Framework 4.8 或兼容的 4.8.1**；单独安装 .NET 9 不能代替它。本次没有添加自动安装运行库、双运行时引导器或在线下载步骤。Windows 7 需 SP1；仍需在真实 Win7/Win10、低配电脑和具体安装程序上验证。

若系统已有 CLR 4、但 Framework 版本不足，入口只读取一次 Framework Release 注册表值并提示缺少 4.8；不扫描运行环境或软件版本。完全没有 CLR 4 的系统由 Windows 的 CLR 装载提示处理。

Framework 与现代 .NET 是不同实现；现代 .NET 的依赖系统运行时部署和自包含部署也是不同模式。参考：[.NET 发布模式](https://learn.microsoft.com/en-us/dotnet/core/deploying/)、[Framework 4.8 的系统兼容性](https://devblogs.microsoft.com/dotnet/announcing-the-net-framework-4-8/)、[单文件自解压说明](https://learn.microsoft.com/en-us/dotnet/core/deploying/single-file/)。

## 实测与边界

测试机是 Windows 11（10.0.26200），不是 Win10/Win7 或低配置电脑。样例使用相同的约 64 MiB ZIP 资源、120 款软件、3,840 个资源引用，不运行安装器。修改前副本只加入同样的测量入口，打包参数与原模板一致。所有测试文件位于 `.build`。

两次使用新文件路径的首启测量如下；“启动”是进程创建到诊断进程退出，进程在界面和清单就绪后立即写报告退出，包含少量关闭开销。没有清空系统磁盘缓存，不能称为严格冷启动。

| 指标 | 修改前 | 修改后 |
| --- | ---: | ---: |
| 首次启动，第 1 组 | 5.259 秒 | 1.390 秒 |
| 首次启动，第 2 组 | 5.085 秒 | 1.210 秒 |
| 第 2 组重复启动，两次平均 | 0.492 秒 | 0.595 秒 |
| 第 2 组首启工作集 | 76.7 MiB | 70.7 MiB |
| 第 2 组首启私有内存 | 28.6 MiB | 48.2 MiB |
| 模板体积 | 109.44 MiB | 约 2.23 MiB |

首启改善明显，但**重复启动略慢，私有内存增加**；不能据此声称所有性能指标都降低。更小的模板减少装载和扫描对象，实际数值仍受磁盘、系统缓存、防护和目标机 Framework 状态影响。

原始记录：`.build/integration-6fd23d72160b469e8f2306206752f5ae/startup-comparison.json`、`.build/integration-fb5636ac47d448669a86d77d9353740a/startup-comparison.json`。样例不含真实第三方安装器；其 EXE 仅用于开发查看清单，不用于安装软件。

## 必要验证

- Release 编译：生成器、测试引用的 net9 执行器和实际 net48 执行器均无警告、无错误。
- 资源读取 9 项：边读边校验、同长度损坏拦截、缺失/长度错误按需拒绝、取消、路径边界、损坏尾部和字节进度。
- 进程控制 11 项：保留已有调用的停止语义，另验证合集取消能终止提前退出的启动器所留下的子进程，并且后续安装器未启动。
- 队列与界面 17 项：顺序准备、跳过/依赖、取消准备、列表与选择，以及关闭窗口先取消而不是立即释放仍在使用的资源。
- 实际 net48 程序集探针 2 项：独立读取清单、只提取选择项并校验摘要；使用无害进程替身确认取消子进程、停止后续队列。探针源码为 `SoftwareInstaller.Tests/FrameworkProbe.cs`，以 `FRAMEWORK_PROBE` 条件编译并引用实际 net48 执行器；不依赖测试端 net9 的运行逻辑。
- 生成逻辑 4 项：接受新模板、拒绝旧模板、实际生成/读取/校验开发样例、确认资源未二次压缩。
- 启动测量每组 12 项确认清单可用；不代表进行 12 次严格冷启动。

用户确认打包后的 v1.2.4 发布前验证：实际发布程序集完成读取 9、进程 11、队列/界面 17、导出流程 13、生成封装 4，加实际 Framework 执行器探针 2，共 56 项定向检查通过。发布构建没有警告或错误；发布包、验证副本和正式 Runner 模板已核对摘要。截图更新自实际发布组件生成的 EXE，不使用旧模板替代。

针对性检查入口（不执行默认完整测试）：

```powershell
dotnet build SoftwareInstaller.Runner -c Release -f net48
dotnet build SoftwareInstaller.Tests -c Release
dotnet run --project SoftwareInstaller.Tests -c Release --no-build -- --bundle-reading-only
dotnet run --project SoftwareInstaller.Tests -c Release --no-build -- --process-control-only
dotnet run --project SoftwareInstaller.Tests -c Release --no-build -- --runner-performance-only
dotnet run --project SoftwareInstaller.Tests -c Release --no-build -- --runner-packaging-only SoftwareInstaller.Runner/bin/Release/net48/SoftwareInstaller.Runner.exe
# 明确提供修改前、后的开发模板路径；会在 .build 创建测量样例，不安装软件。
dotnet run --project SoftwareInstaller.Tests -c Release --no-build -- --runner-startup-only <before.exe> <after.exe>
```

v1.2.4 按用户后续明确要求构建发布模板/便携包。已有合集不会自动变化，需要用新模板重新生成；使用旧模板会在生成检查时被拒绝。
