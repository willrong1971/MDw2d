# 第三方组件

本地原型包含 CPython 3.12 和 LibreOfficeDev。转换全部在本机执行。

## CPython

CPython 由 Python Software Foundation 及其贡献者提供，遵循 Python Software Foundation License 及其随附的其他许可。打包时保留 `runtime/python/lib/python3.12/LICENSE.txt`、运行时 `share` 目录，以及标准库自带的许可与说明。

项目和许可来源：
- https://www.python.org/
- https://docs.python.org/3/license.html

## LibreOffice

LibreOffice 由 The Document Foundation 及其贡献者提供，其主要许可为 Mozilla Public License 2.0，并包含采用其他许可证的组件。打包时保留完整原始 LibreOfficeDev.app，包括上游许可、版权和第三方声明。具体许可请查看 `libreoffice/LibreOfficeDev.app/Contents/Resources` 及程序随附的许可文件。

项目和许可来源：
- https://www.libreoffice.org/
- https://www.libreoffice.org/about-us/licenses/

这是本地开发原型，尚未作为正式软件发行。对外分发前，应完整核查所有随附许可和源代码获取要求，并使用开发者签名及公证。
