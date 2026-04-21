# 货盘收集与比价 MVP

这是一个从 0 搭建的轻量版系统，用来解决两个核心问题：

1. 生成一个供应商上传链接，发到微信群里收集货盘
2. 统一解析 Excel / CSV 后，在后台做基础比价

## 功能

- 创建“收集批次”，自动生成供应商上传链接
- 供应商通过公开链接上传 `.xlsx`、`.xls`、`.csv`
- 自动解析首个工作表并入库到 SQLite
- 后台查看每个批次的上传记录
- 提供比价接口和后台比价页面
- 提供一个可下载的 Excel 模板

## 技术栈

- Node.js
- Express
- SQLite (`better-sqlite3`)
- Excel 解析 (`xlsx`)

## 启动

```bash
npm install
npm run dev
```

然后打开：

- 后台管理页：[http://localhost:3000/admin.html](http://localhost:3000/admin.html)
- 供应商模板下载：[http://localhost:3000/template.xlsx](http://localhost:3000/template.xlsx)

## API

### 1. 创建收集批次

`POST /api/collections`

请求体：

```json
{
  "name": "4月水果供应商收集",
  "description": "请上传最新供货价、起订量和库存"
}
```

### 2. 获取批次列表

`GET /api/collections`

### 3. 获取批次上传记录

`GET /api/collections/:id/uploads`

### 4. 比价接口

`GET /api/collections/:id/compare?keyword=苹果&groupBy=smart`

支持的 `groupBy`：

- `smart`：优先按 SKU，否则按 商品名 + 规格
- `sku`：优先按 SKU
- `name`：按 商品名 + 规格

### 5. 供应商上传

`POST /api/public/upload/:token`

表单字段：

- `supplierName`
- `file`

## 默认识别的表头

系统会自动识别这类常见中文 / 英文列名：

- 商品名称：`商品名称`、`货品名称`、`品名`、`productName`
- SKU：`SKU`、`货号`、`条码`
- 规格：`规格`、`规格型号`、`spec`
- 价格：`供货价`、`采购价`、`单价`、`price`
- 单位：`单位`、`计价单位`
- 起订量：`起订量`、`最小起订量`、`MOQ`
- 库存：`库存`、`可供数量`

## 下一步建议

这个版本先解决“收”和“比”两个最关键环节。接下来很适合继续加：

- 登录和权限控制
- 一个供应商多次上传时的版本管理
- 更强的商品匹配规则
- 审核字段映射，让每家供应商模板都能兼容
- 对外开放正式的比价 API 鉴权
