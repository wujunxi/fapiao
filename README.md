# PDF发票抽取到Excel文件

将发票中的关键信息抽取到excel文件，需要对发票中的文本框进行定位。
## 环境准备

- Nodejs
  访问官网，下载安装对应操作系统的 Node.js https://nodejs.org/en/

## 下载脚本

- 使用git下载

    ```sh
    git clone git@github.com:wujunxi/fapiao.git
    ```

- 使用压缩包下载，点击 code -> Download ZIP，下载后解压。

## Usage

进入脚本目录，先安装依赖 `npm install`，然后执行 `node index.js mydir`，其中mydir替换为发票pdf所在目录.示例：
```sh
node index.js 'data/2022-01至2022-05/2022-01至2022-05'
```
结果将会保存到当前目录下 result.xlsx 文件

## 发票类型适配
- [x] 广东增值税电子普通发票  
- [ ] 深圳增值税电子普通发票  
- [ ] 浙江增值税电子普通发票  
- [ ] 上海增值税电子普通发票  
- [ ] 河南增值税电子普通发票  
- [ ] 北京增值税电子普通发票  
- [ ] 湖北增值税电子普通发票  
- [ ] 天津增值税电子普通发票  
- [ ] 深圳电子普通发票  
- [ ] 未知  