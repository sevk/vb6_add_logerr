# VB6 Add LogErr 插件

这是一个用于 VB6 的插件，旨在帮助开发者在开发过程中更方便地添加日志记录功能。

## 功能简介

- 提供日志记录功能，便于调试和错误追踪。
- 与 VB6 集成，提供便捷的插件式体验。

## 文件说明

- `AddIn.bas`: 插件的主要功能实现文件。
- `Connect.cls`: 插件连接和初始化相关类文件。
- `MenuHandler.cls`: 菜单处理类文件，用于在 VB6 中创建和处理插件菜单。
- `add_logerr.dll`: 编译后的插件动态链接库文件。
- `add_logerr.vbp`: VB6 插件项目文件。
- `add_logerr.vbw`: VB6 插件工作区文件。
- `.gitignore`: 指定 Git 版本控制系统忽略的文件和目录。
- `LICENSE`: 项目使用的开源许可证文件。

## 安装与使用

1. 将 `add_logerr.dll` 文件复制到 VB6 的插件目录中, 或直接运行cmd , regsvr32 add_logerr.dll 
2. 在 VB6 中加载插件。
3. 使用插件提供的功能在代码中添加日志记录。

## 许可证

本项目基于 MIT 许可证。详情请查看 LICENSE 文件。

## 贡献

欢迎贡献代码和报告问题。请通过项目的 Git 仓库提交 Pull Request 或 Issue。