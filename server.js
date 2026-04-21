const express = require("express");
const multer = require("multer");
const Database = require("better-sqlite3");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const app = express();
const rootDir = __dirname;
loadLocalEnvFiles([
  path.join(rootDir, ".env.local"),
  path.join(rootDir, ".env")
]);
const publicDir = path.join(rootDir, "public");
const dataDir = path.join(rootDir, "data");
const tempUploadDir = path.join(dataDir, "uploads_tmp");
const catalogUploadDir = path.join(dataDir, "catalog_uploads");
const legacyUploadDir = path.join(dataDir, "uploads");
const dbPath = path.join(dataDir, "app.db");
const port = Number(process.env.PORT || 3002);
const maxUploadSize = 50 * 1024 * 1024;
const defaultAiProvider = "dashscope-compatible";
const defaultAiBaseUrl = process.env.DASHSCOPE_BASE_URL || "https://dashscope.aliyuncs.com/compatible-mode/v1";
const defaultAiModel = process.env.DASHSCOPE_MODEL || "qwen-plus-latest";
const defaultQjlAuthBase = process.env.QJL_AUTH_BASE || "http://81.68.227.162:3367";
const defaultQjlApiBase = process.env.QJL_API_BASE || "https://apipro.qunjielong.com";
const defaultOneboundApiBase = process.env.ONEBOUND_API_BASE || "https://api-gw.onebound.cn";
const defaultOneboundApiKey = process.env.ONEBOUND_API_KEY || "";
const defaultOneboundApiSecret = process.env.ONEBOUND_API_SECRET || "";

const qjlBaseHeaders = {
  "feature-tag": "f0000",
  "device-type": "5",
  "client-version": "6.2.12",
  "content-type": "application/json",
  appid: "wx059cd327295ab444",
  "User-Agent":
    "Mozilla/5.0 (iPhone; CPU iPhone OS 15_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.0 Mobile/15E148 Safari/604.1 wechatdevtools/1.06.2307260 MicroMessenger/8.0.5 Language/zh_CN webview/",
  sceneCode: "1001",
  companyId: "190",
  Accept: "*/*",
  Referer: "https://servicewechat.com/wx059cd327295ab444/devtools/page-frame.html"
};

const qjlMiniRoutes = {
  homepage: "pro/pages/homepage/group-homepage/group-homepage",
  homepageList: "pro/pages/homepage/group-homepage/group-homepage",
  dataOverview: "pro/pages/data-analyse/data-analyse-v2-overview-page/data-analyse-v2-overview-page"
};

const externalCompareSources = {
  "1688": {
    path: "/1688/item_search/",
    label: "1688"
  },
  jd: {
    path: "/jd/item_search/",
    label: "京东"
  },
  taobao: {
    path: "/taobao/item_search/",
    label: "淘宝"
  }
};

const qjlProfileCategoryKeywords = {
  服务课程: ["课程", "私教", "训练", "预约", "体验课", "游泳", "瑜伽", "训练营", "服务"],
  美妆护肤: ["精华", "面膜", "面霜", "防晒", "口红", "唇釉", "眼霜", "气垫", "护肤"],
  香氛个护: ["香氛", "香水", "洗发", "护发", "沐浴", "身体乳", "牙膏", "漱口水"],
  茶饮滋补: ["燕窝", "阿胶", "花胶", "茶", "红茶", "绿茶", "滋补", "养生"],
  服饰鞋包: ["短袖", "衬衫", "内裤", "文胸", "袜", "鞋", "包", "睡衣", "裤"],
  水果生鲜: ["水果", "橙", "苹果", "蓝莓", "榴莲", "牛油果", "玉米", "鸡蛋", "车厘子"],
  零食饮品: ["坚果", "牛奶", "酸奶", "咖啡", "零食", "蜂蜜", "饼干", "饮料", "面包"],
  肉禽海鲜: ["牛肉", "鸡", "鸭", "鱼", "虾", "海鲜", "排骨", "羊肉", "鸡煲"],
  数码充值: ["充电宝", "快充", "充值", "话费", "流量", "耳机", "数据线"],
  家居日用: ["纸巾", "湿巾", "收纳", "拖把", "厨房", "床品", "清洁", "毛巾"],
  母婴用品: ["宝宝", "婴儿", "儿童", "纸尿裤", "奶瓶", "童装", "辅食"],
  保健个护: ["保健", "益生菌", "鱼油", "维生素", "胶原", "喷雾", "贴", "护具"]
};

const aiRuntimeConfig = {
  provider: defaultAiProvider,
  apiKey: process.env.DASHSCOPE_API_KEY || "",
  baseUrl: defaultAiBaseUrl,
  model: defaultAiModel,
  source: process.env.DASHSCOPE_API_KEY ? "environment" : "missing"
};

ensureDir(dataDir);
ensureDir(tempUploadDir);
ensureDir(catalogUploadDir);

const db = new Database(dbPath);
db.pragma("journal_mode = WAL");
initDb();
loadAiConfig();

const upload = multer({
  storage: multer.diskStorage({
    destination: (_req, _file, cb) => cb(null, tempUploadDir),
    filename: (_req, file, cb) => {
      const extension = path.extname(file.originalname || "").toLowerCase() || ".xlsx";
      cb(null, `${crypto.randomUUID()}${extension}`);
    }
  }),
  limits: {
    fileSize: maxUploadSize
  },
  fileFilter: (_req, file, cb) => {
    const extension = path.extname(file.originalname || "").toLowerCase();
    const allowed = new Set([".xlsx", ".xls", ".csv"]);
    if (!allowed.has(extension)) {
      cb(new Error("仅支持 .xlsx、.xls、.csv 文件"));
      return;
    }

    cb(null, true);
  }
});

const headerAliases = {
  productName: [
    "商品名称",
    "货品名称",
    "产品名称",
    "商品信息",
    "货品信息",
    "商品标题",
    "标题",
    "款名",
    "宝贝名称",
    "品名",
    "名称",
    "商品",
    "货品",
    "productname",
    "product",
    "name"
  ],
  sku: [
    "sku",
    "skuid",
    "商品编码",
    "商品编号",
    "货品编号",
    "货号",
    "款号",
    "款式编码",
    "商家编码",
    "条码",
    "编码",
    "规格编码"
  ],
  spec: [
    "规格",
    "规格型号",
    "规格名称",
    "规格信息",
    "颜色尺码",
    "型号",
    "包装规格",
    "单位规格",
    "spec"
  ],
  price: [
    "价格",
    "供货价",
    "进货价",
    "采购价",
    "批发价",
    "档口价",
    "拿货价",
    "结算价",
    "团购价",
    "店主价",
    "供货单价",
    "单价",
    "售价",
    "price"
  ],
  unit: [
    "单位",
    "计价单位",
    "销售单位",
    "包装单位",
    "unit"
  ],
  moq: [
    "起订量",
    "最小起订量",
    "最小订购量",
    "最低起订量",
    "起批量",
    "moq"
  ],
  stock: [
    "库存",
    "可供数量",
    "可售库存",
    "库存量",
    "stock"
  ],
  category: [
    "品类",
    "商品品类",
    "类目",
    "商品类目",
    "分类",
    "商品分类",
    "货品分类",
    "一级类目",
    "二级类目",
    "三级类目",
    "category",
    "categoryname"
  ],
  color: [
    "颜色",
    "颜色分类",
    "颜色规格"
  ],
  size: [
    "尺码",
    "尺寸",
    "码数",
    "规格尺码"
  ]
};

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(publicDir));

app.get("/", (_req, res) => {
  res.redirect("/admin.html");
});

app.get("/health", (_req, res) => {
  res.json({ ok: true, now: new Date().toISOString() });
});

app.get("/template.xlsx", (_req, res) => {
  const rows = [
    {
      商品名称: "山东苹果",
      SKU: "APL-001",
      规格: "80果 / 5kg",
      供货价: 26.8,
      单位: "箱",
      起订量: 10,
      库存: 320
    },
    {
      商品名称: "麒麟西瓜",
      SKU: "XG-002",
      规格: "6kg 左右",
      供货价: 18.5,
      单位: "个",
      起订量: 20,
      库存: 100
    }
  ];

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, worksheet, "货盘模板");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader("Content-Disposition", "attachment; filename=\"supplier-template.xlsx\"");
  res.send(buffer);
});

app.get("/u/:token", (_req, res) => {
  res.status(410).send(`
    <!DOCTYPE html>
    <html lang="zh-CN">
      <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>上传入口已停用</title>
        <style>
          body { margin: 0; min-height: 100vh; display: grid; place-items: center; background: #f7f1e5; font-family: "PingFang SC", "Noto Sans SC", "Microsoft YaHei", sans-serif; color: #322d25; }
          .card { width: min(560px, calc(100% - 40px)); background: white; border-radius: 24px; padding: 28px; box-shadow: 0 16px 60px rgba(50, 45, 37, 0.12); }
          h1 { margin-top: 0; }
          p { line-height: 1.7; color: #6b5f52; }
        </style>
      </head>
      <body>
        <div class="card">
          <h1>旧的上传链接模式已停用</h1>
          <p>请在软件首页创建收集批次时，直接选择并上传多个货盘文件。系统会自动把这些文件归入同一个批次。</p>
        </div>
      </body>
    </html>
  `);
});

app.post("/api/collections", (req, res) => {
  const name = String(req.body.name || "").trim();
  const description = String(req.body.description || "").trim();

  if (!name) {
    res.status(400).json({ error: "请填写收集批次名称" });
    return;
  }

  const id = crypto.randomUUID();
  const uploadToken = crypto.randomBytes(18).toString("hex");
  const now = new Date().toISOString();

  db.prepare(
    `
    INSERT INTO collections (id, name, description, upload_token, created_at)
    VALUES (?, ?, ?, ?, ?)
    `
  ).run(id, name, description, uploadToken, now);

  res.status(201).json(formatCollection(getCollectionById(id)));
});

app.post("/api/collections/import-batch", upload.array("files", 30), async (req, res) => {
  const name = String(req.body.name || "").trim();
  const description = String(req.body.description || "").trim();
  const files = Array.isArray(req.files) ? req.files : [];

  if (!name) {
    cleanupUploads(files);
    res.status(400).json({ error: "请填写收集批次名称" });
    return;
  }

  if (!files.length) {
    res.status(400).json({ error: "请至少上传一个 Excel 或 CSV 文件" });
    return;
  }

  const collection = {
    id: crypto.randomUUID(),
    name,
    description,
    upload_token: crypto.randomBytes(18).toString("hex"),
    created_at: new Date().toISOString()
  };

  try {
    const parsedFiles = [];

    for (const file of files) {
      const supplierName = supplierNameFromFilename(file.originalname);
      const parsed = parseWorkbook(file.path, supplierName);

      if (parsed.parsedCount === 0) {
        throw new Error(`文件《${file.originalname}》没有识别到可导入的商品数据`);
      }

      const storedFile = moveUploadIntoCollectionFolder({
        collectionId: collection.id,
        file
      });
      const summary = await summarizeCatalogForStorage({
        collection,
        supplierName,
        originalFilename: file.originalname,
        parsed
      });

      parsedFiles.push({
        file,
        storedFile,
        supplierName,
        parsed,
        summary
      });
    }

    const result = importBatchFilesIntoCollection({ collection, parsedFiles });

    res.status(201).json({
      ...formatCollection(getCollectionById(collection.id)),
      fileCount: parsedFiles.length,
      totalRowCount: result.totalRowCount,
      totalParsedCount: result.totalParsedCount
    });
  } catch (error) {
    cleanupUploads(files);
    cleanupStoredUploadsByCollection(collection.id);
    res.status(400).json({ error: error.message || "批量导入失败，请检查文件格式后重试" });
  }
});

app.get("/api/collections", (_req, res) => {
  const rows = db
    .prepare(
      `
      SELECT
        c.id,
        c.name,
        c.description,
        c.upload_token,
        c.created_at,
        COUNT(DISTINCT su.id) AS upload_count,
        COUNT(DISTINCT p.id) AS product_count
      FROM collections c
      LEFT JOIN supplier_uploads su ON su.collection_id = c.id
      LEFT JOIN products p ON p.collection_id = c.id
      GROUP BY c.id
      ORDER BY c.created_at DESC
      `
    )
    .all();

  res.json({ items: rows.map(formatCollection) });
});

app.get("/api/collections/:id", (req, res) => {
  const collection = getCollectionById(req.params.id);

  if (!collection) {
    res.status(404).json({ error: "未找到该收集批次" });
    return;
  }

  res.json(formatCollection(collection));
});

app.get("/api/collections/:id/uploads", (req, res) => {
  const collection = getCollectionById(req.params.id);

  if (!collection) {
    res.status(404).json({ error: "未找到该收集批次" });
    return;
  }

  const uploads = db
    .prepare(
      `
      SELECT
        id,
        supplier_name,
        original_filename,
        row_count,
        parsed_count,
        uploaded_at
      FROM supplier_uploads
      WHERE collection_id = ?
      ORDER BY uploaded_at DESC
      `
    )
    .all(req.params.id);

  res.json({
    collection: formatCollection(collection),
    items: uploads
  });
});

app.get("/api/collections/:id/upload-batches", (req, res) => {
  const collection = getCollectionById(req.params.id);

  if (!collection) {
    res.status(404).json({ error: "未找到该收集批次" });
    return;
  }

  const uploads = db
    .prepare(
      `
      SELECT
        id,
        supplier_name,
        original_filename,
        stored_filename,
        parsed_count,
        category_summary,
        product_names_summary,
        catalog_overview,
        analysis_source,
        uploaded_at
      FROM supplier_uploads
      WHERE collection_id = ?
      ORDER BY uploaded_at DESC, id DESC
      `
    )
    .all(req.params.id);

  const productRows = db
    .prepare(
      `
      SELECT
        upload_id,
        product_name,
        raw_json
      FROM products
      WHERE collection_id = ?
      ORDER BY id ASC
      `
    )
    .all(req.params.id);

  const productsByUpload = new Map();
  for (const row of productRows) {
    if (!productsByUpload.has(row.upload_id)) {
      productsByUpload.set(row.upload_id, []);
    }
    productsByUpload.get(row.upload_id).push(row);
  }

  const items = uploads.map((upload) =>
    buildUploadBatchItem({
      collection,
      upload,
      products: productsByUpload.get(upload.id) || []
    })
  );

  res.json({
    collection: formatCollection(collection),
    total: items.length,
    items
  });
});

app.get("/api/upload-batches", (_req, res) => {
  const uploads = db
    .prepare(
      `
      SELECT
        su.id,
        su.collection_id,
        su.supplier_name,
        su.original_filename,
        su.stored_filename,
        su.parsed_count,
        su.category_summary,
        su.product_names_summary,
        su.catalog_overview,
        su.analysis_source,
        su.uploaded_at,
        c.name AS collection_name,
        c.created_at AS collection_created_at
      FROM supplier_uploads su
      INNER JOIN collections c ON c.id = su.collection_id
      ORDER BY su.uploaded_at DESC, su.id DESC
      `
    )
    .all();

  const productRows = db
    .prepare(
      `
      SELECT
        upload_id,
        product_name,
        raw_json
      FROM products
      ORDER BY id ASC
      `
    )
    .all();

  const productsByUpload = new Map();
  for (const row of productRows) {
    if (!productsByUpload.has(row.upload_id)) {
      productsByUpload.set(row.upload_id, []);
    }
    productsByUpload.get(row.upload_id).push(row);
  }

  const items = uploads.map((upload) =>
    buildUploadBatchItem({
      collection: {
        name: upload.collection_name,
        created_at: upload.collection_created_at
      },
      upload,
      products: productsByUpload.get(upload.id) || []
    })
  );

  res.json({
    total: items.length,
    items
  });
});

app.get("/api/uploads/:id/file", (req, res) => {
  const upload = db
    .prepare(
      `
      SELECT
        id,
        collection_id,
        original_filename,
        stored_filename
      FROM supplier_uploads
      WHERE id = ?
      `
    )
    .get(req.params.id);

  if (!upload) {
    res.status(404).json({ error: "未找到这份货盘文件" });
    return;
  }

  const filePath = resolveStoredUploadPath(upload);
  if (!filePath || !fs.existsSync(filePath)) {
    res.status(404).json({ error: "原始货盘文件不存在，可能已被移动或清理" });
    return;
  }

  res.setHeader("Content-Disposition", `inline; filename*=UTF-8''${encodeURIComponent(upload.original_filename)}`);
  res.sendFile(filePath);
});

app.get("/api/ai/session", (_req, res) => {
  res.json({
    enabled: Boolean(aiRuntimeConfig.apiKey),
    provider: aiRuntimeConfig.provider,
    model: aiRuntimeConfig.model,
    baseUrl: aiRuntimeConfig.baseUrl,
    source: aiRuntimeConfig.source
  });
});

app.get("/api/llm/config", (_req, res) => {
  res.json({
    enabled: Boolean(aiRuntimeConfig.apiKey),
    provider: aiRuntimeConfig.provider,
    model: aiRuntimeConfig.model,
    baseUrl: aiRuntimeConfig.baseUrl,
    source: aiRuntimeConfig.source,
    hasApiKey: Boolean(aiRuntimeConfig.apiKey),
    apiKeyHint: maskSecret(aiRuntimeConfig.apiKey)
  });
});

app.post("/api/llm/config", (req, res) => {
  const provider = String(req.body.provider || "").trim() || defaultAiProvider;
  const model = String(req.body.model || "").trim() || defaultAiModel;
  const baseUrl = String(req.body.baseUrl || "").trim() || defaultAiBaseUrl;
  const apiKey = String(req.body.apiKey || "").trim();

  if (!model) {
    res.status(400).json({ error: "请填写模型名称" });
    return;
  }

  if (!baseUrl) {
    res.status(400).json({ error: "请填写接口地址" });
    return;
  }

  const nextApiKey = apiKey || aiRuntimeConfig.apiKey;
  if (!nextApiKey) {
    res.status(400).json({ error: "请填写 API Key" });
    return;
  }

  persistSetting("llm_provider", provider);
  persistSetting("llm_model", model);
  persistSetting("llm_base_url", baseUrl.replace(/\/$/, ""));
  persistSetting("llm_api_key", nextApiKey);
  loadAiConfig();

  res.json({
    ok: true,
    enabled: Boolean(aiRuntimeConfig.apiKey),
    provider: aiRuntimeConfig.provider,
    model: aiRuntimeConfig.model,
    baseUrl: aiRuntimeConfig.baseUrl,
    source: aiRuntimeConfig.source,
    hasApiKey: Boolean(aiRuntimeConfig.apiKey),
    apiKeyHint: maskSecret(aiRuntimeConfig.apiKey)
  });
});

app.delete("/api/llm/config", (_req, res) => {
  deleteSetting("llm_provider");
  deleteSetting("llm_model");
  deleteSetting("llm_base_url");
  deleteSetting("llm_api_key");
  loadAiConfig();

  res.json({
    ok: true,
    enabled: Boolean(aiRuntimeConfig.apiKey),
    provider: aiRuntimeConfig.provider,
    model: aiRuntimeConfig.model,
    baseUrl: aiRuntimeConfig.baseUrl,
    source: aiRuntimeConfig.source,
    hasApiKey: Boolean(aiRuntimeConfig.apiKey),
    apiKeyHint: maskSecret(aiRuntimeConfig.apiKey)
  });
});

app.get("/api/qjl/account", (_req, res) => {
  const pendingLogin = getPendingQjlLogin();
  const account = getCurrentQjlAccount();

  res.json({
    ok: true,
    loggedIn: Boolean(account?.uid && account?.mini_token),
    pendingLogin: pendingLogin.pending,
    loginUrl: pendingLogin.pending ? pendingLogin.loginUrl : "",
    account: formatQjlAccount(account)
  });
});

app.post("/api/qjl/auth/login/start", async (_req, res) => {
  try {
    const login = await requestQjlLoginKey();
    persistPendingQjlLogin(login);

    res.json({
      ok: true,
      pendingLogin: true,
      key: login.key,
      loginUrl: login.loginUrl
    });
  } catch (error) {
    res.status(500).json({ error: error.message || "获取群接龙登录链接失败" });
  }
});

app.post("/api/qjl/auth/poll", async (req, res) => {
  try {
    const pending = getPendingQjlLogin();
    const key = String(req.body.key || pending.key || "").trim();

    if (!key) {
      res.status(400).json({ error: "当前没有待确认的群接龙登录，请先点击开始登录。" });
      return;
    }

    const loginResult = await pollQjlLoginByKey(key);
    if (!loginResult.miniToken) {
      res.json({
        ok: true,
        loggedIn: false,
        pendingLogin: true,
        loginUrl: pending.loginUrl || "",
        status: "pending",
        message: "还没有检测到登录结果，请在群接龙完成登录后再检查一次。"
      });
      return;
    }

    const account = await buildQjlAccountFromLogin(loginResult);
    persistQjlAccount(account);
    clearPendingQjlLogin();

    res.json({
      ok: true,
      loggedIn: true,
      pendingLogin: false,
      account: formatQjlAccount(getQjlAccountByUid(account.uid)),
      message: "群接龙已登录，并已生成用户画像。"
    });
  } catch (error) {
    res.status(500).json({ error: error.message || "检查群接龙登录状态失败" });
  }
});

app.post("/api/qjl/profile/refresh", async (_req, res) => {
  try {
    const current = getCurrentQjlAccount();
    if (!current?.uid || !current?.mini_token) {
      res.status(400).json({ error: "当前还没有登录群接龙，无法刷新用户画像。" });
      return;
    }

    const account = await buildQjlAccountFromLogin({
      uid: String(current.uid),
      miniToken: String(current.mini_token)
    });

    persistQjlAccount({
      ...account,
      createdAt: current.created_at || account.createdAt
    });

    res.json({
      ok: true,
      loggedIn: true,
      account: formatQjlAccount(getQjlAccountByUid(account.uid)),
      message: "用户画像已刷新。"
    });
  } catch (error) {
    res.status(500).json({ error: error.message || "刷新用户画像失败" });
  }
});

app.post("/api/qjl/homepage/switch", async (req, res) => {
  try {
    const current = getCurrentQjlAccount();
    if (!current?.uid || !current?.mini_token) {
      res.status(400).json({ error: "当前还没有登录群接龙，无法切换主页。" });
      return;
    }

    const ghCode = String(req.body.ghCode || "").trim();
    if (!ghCode) {
      res.status(400).json({ error: "请先选择要切换的群接龙主页。" });
      return;
    }

    const account = await rebuildQjlAccountForHomepage({
      account: current,
      ghCode
    });

    persistQjlAccount(account);

    res.json({
      ok: true,
      loggedIn: true,
      account: formatQjlAccount(getQjlAccountByUid(account.uid)),
      message: `已切换到主页：${account.ghName || ghCode}`
    });
  } catch (error) {
    res.status(500).json({ error: error.message || "切换群接龙主页失败" });
  }
});

app.delete("/api/qjl/account", (_req, res) => {
  deleteCurrentQjlAccount();

  res.json({
    ok: true,
    loggedIn: false,
    pendingLogin: false,
    account: null
  });
});

app.get("/api/collections/:id/products", (req, res) => {
  const collection = getCollectionById(req.params.id);

  if (!collection) {
    res.status(404).json({ error: "未找到该收集批次" });
    return;
  }

  const keyword = String(req.query.keyword || "").trim();

  let query = `
    SELECT
      id,
      supplier_name,
      product_name,
      sku,
      spec,
      unit,
      price,
      moq,
      stock,
      uploaded_at
    FROM products
    WHERE collection_id = ?
  `;
  const params = [req.params.id];

  if (keyword) {
    query += `
      AND (
        product_name LIKE ?
        OR sku LIKE ?
        OR spec LIKE ?
        OR supplier_name LIKE ?
      )
    `;
    const fuzzy = `%${keyword}%`;
    params.push(fuzzy, fuzzy, fuzzy, fuzzy);
  }

  query += " ORDER BY uploaded_at DESC, supplier_name ASC, product_name ASC, id DESC";

  const items = db.prepare(query).all(...params);

  res.json({
    collection: formatCollection(collection),
    keyword,
    total: items.length,
    items
  });
});

app.get("/api/collections/:id/organized-products", async (req, res) => {
  const collection = getCollectionById(req.params.id);

  if (!collection) {
    res.status(404).json({ error: "未找到该收集批次" });
    return;
  }

  if (!aiRuntimeConfig.apiKey) {
    res.status(400).json({
      error: "还没有启用百炼 AI。请先在后台页面填入 API Key，再使用 AI 整理货盘。"
    });
    return;
  }

  const keyword = String(req.query.keyword || "").trim();

  let query = `
    SELECT
      id,
      supplier_name,
      product_name,
      sku,
      spec,
      unit,
      price,
      moq,
      stock,
      uploaded_at,
      raw_json
    FROM products
    WHERE collection_id = ?
  `;
  const params = [req.params.id];

  if (keyword) {
    query += `
      AND (
        product_name LIKE ?
        OR sku LIKE ?
        OR spec LIKE ?
        OR supplier_name LIKE ?
      )
    `;
    const fuzzy = `%${keyword}%`;
    params.push(fuzzy, fuzzy, fuzzy, fuzzy);
  }

  query += " ORDER BY uploaded_at DESC, supplier_name ASC, id DESC";

  const sourceRows = db.prepare(query).all(...params).map((row) => ({
    id: row.id,
    supplierName: row.supplier_name,
    uploadedAt: row.uploaded_at,
    raw: compactRawRow(
      safeParseJson(row.raw_json) || {
        商品名称: row.product_name,
        SKU: row.sku,
        规格: row.spec,
        单位: row.unit,
        价格: row.price,
        起订量: row.moq,
        库存: row.stock
      }
    )
  }));

  const organized = await organizeProductsWithAi({
    collection,
    keyword,
    sourceRows
  });

  res.json(organized);
});

app.get("/api/collections/:id/analysis", async (req, res) => {
  const collection = getCollectionById(req.params.id);

  if (!collection) {
    res.status(404).json({ error: "未找到该收集批次" });
    return;
  }

  if (!aiRuntimeConfig.apiKey) {
    res.status(400).json({ error: "请先配置大模型，再使用 AI 汇总分析。" });
    return;
  }

  const rows = db
    .prepare(
      `
      SELECT
        supplier_name,
        product_name,
        sku,
        spec,
        unit,
        price,
        moq,
        stock,
        uploaded_at
      FROM products
      WHERE collection_id = ?
      ORDER BY uploaded_at DESC, supplier_name ASC, id DESC
      `
    )
    .all(req.params.id);

  const analysis = await summarizeCollectionWithAi({ collection, rows });
  res.json(analysis);
});

app.post("/api/collections/:id/upload-chat", async (req, res) => {
  const collection = getCollectionById(req.params.id);

  if (!collection) {
    res.status(404).json({ error: "未找到该收集批次" });
    return;
  }

  if (!aiRuntimeConfig.apiKey) {
    res.status(400).json({ error: "请先配置大模型，再使用对话分析。" });
    return;
  }

  const message = String(req.body.message || "").trim();
  const uploadIds = Array.isArray(req.body.uploadIds)
    ? req.body.uploadIds.map((item) => String(item || "").trim()).filter(Boolean)
    : [];
  const history = Array.isArray(req.body.history) ? req.body.history : [];

  if (!uploadIds.length) {
    res.status(400).json({ error: "请先勾选至少一份货盘，再发起对话。" });
    return;
  }

  if (!message) {
    res.status(400).json({ error: "请输入你想让 AI 分析的问题。" });
    return;
  }

  const selectedUploads = getUploadsByIds(req.params.id, uploadIds);
  if (!selectedUploads.length) {
    res.status(400).json({ error: "没有找到你勾选的货盘，请刷新后重试。" });
    return;
  }

  const normalizedIds = selectedUploads.map((item) => item.id);
  const products = getProductsByUploadIds(req.params.id, normalizedIds);
  const context = buildUploadChatContext({
    collection,
    uploads: selectedUploads,
    products
  });

  const answer = await chatAboutUploadsWithAi({
    collection,
    message,
    history,
    context
  });

  res.json({
    ok: true,
    model: aiRuntimeConfig.model,
    selectedCount: selectedUploads.length,
    selectedUploads: selectedUploads.map((item) => ({
      id: item.id,
      supplierName: item.supplier_name,
      catalogName: item.original_filename
    })),
    answer
  });
});

app.post("/api/upload-chat", async (req, res) => {
  if (!aiRuntimeConfig.apiKey) {
    res.status(400).json({ error: "请先配置大模型，再使用对话分析。" });
    return;
  }

  const message = String(req.body.message || "").trim();
  const uploadIds = Array.isArray(req.body.uploadIds)
    ? req.body.uploadIds.map((item) => String(item || "").trim()).filter(Boolean)
    : [];
  const history = Array.isArray(req.body.history) ? req.body.history : [];

  if (!uploadIds.length) {
    res.status(400).json({ error: "请先勾选至少一份货盘，再发起对话。" });
    return;
  }

  if (!message) {
    res.status(400).json({ error: "请输入你想让 AI 分析的问题。" });
    return;
  }

  const selectedUploads = getUploadsByIdsAcrossCollections(uploadIds);
  if (!selectedUploads.length) {
    res.status(400).json({ error: "没有找到你勾选的货盘，请刷新后重试。" });
    return;
  }

  const normalizedIds = selectedUploads.map((item) => item.id);
  const products = getProductsByUploadIdsAcrossCollections(normalizedIds);
  const context = buildUploadChatContextAcrossCollections({
    uploads: selectedUploads,
    products
  });

  const answer = await chatAboutUploadsWithAi({
    collection: {
      name: "历史货盘批次",
      description: "用户从历史批次中勾选了多份货盘"
    },
    message,
    history,
    context
  });

  res.json({
    ok: true,
    model: aiRuntimeConfig.model,
    selectedCount: selectedUploads.length,
    selectedUploads: selectedUploads.map((item) => ({
      id: item.id,
      supplierName: item.supplier_name,
      catalogName: item.original_filename,
      batchName: item.collection_name
    })),
    answer
  });
});

app.get("/api/collections/:id/compare", (req, res) => {
  const collection = getCollectionById(req.params.id);

  if (!collection) {
    res.status(404).json({ error: "未找到该收集批次" });
    return;
  }

  const keyword = String(req.query.keyword || "").trim();
  const groupBy = ["smart", "sku", "name"].includes(req.query.groupBy) ? req.query.groupBy : "smart";

  let query = `
    SELECT
      product_name,
      sku,
      spec,
      unit,
      price,
      moq,
      stock,
      supplier_name,
      uploaded_at
    FROM products
    WHERE collection_id = ?
      AND price IS NOT NULL
  `;
  const params = [req.params.id];

  if (keyword) {
    query += `
      AND (
        product_name LIKE ?
        OR sku LIKE ?
        OR spec LIKE ?
        OR supplier_name LIKE ?
      )
    `;
    const fuzzy = `%${keyword}%`;
    params.push(fuzzy, fuzzy, fuzzy, fuzzy);
  }

  query += " ORDER BY product_name ASC, price ASC";

  const products = db.prepare(query).all(...params);
  const groups = buildComparisonGroups(products, groupBy);

  res.json({
    collection: formatCollection(collection),
    keyword,
    groupBy,
    totalGroups: groups.length,
    items: groups
  });
});

app.get("/api/compare/external-search", async (req, res) => {
  try {
    const source = String(req.query.source || "1688").trim().toLowerCase();
    const matchMode = String(req.query.matchMode || "brand_style").trim().toLowerCase();
    const sourceConfig = externalCompareSources[source];
    if (!sourceConfig) {
      res.status(400).json({ error: "暂不支持这个外部平台。" });
      return;
    }

    const keyword = String(req.query.keyword || "").trim();
    if (!keyword) {
      res.status(400).json({ error: "请输入要搜索的商品关键词。" });
      return;
    }

    if (!defaultOneboundApiKey || !defaultOneboundApiSecret) {
      res.status(400).json({ error: "尚未配置 OneBound 接口凭据，请先设置 ONEBOUND_API_KEY 和 ONEBOUND_API_SECRET。" });
      return;
    }

    const page = Math.max(1, Number.parseInt(String(req.query.page || "1"), 10) || 1);
    const pageSize = Math.min(40, Math.max(1, Number.parseInt(String(req.query.pageSize || "20"), 10) || 20));
    const startPrice = Math.max(0, Number.parseFloat(String(req.query.startPrice || "0")) || 0);
    const endPrice = Math.max(0, Number.parseFloat(String(req.query.endPrice || "0")) || 0);

    const params = new URLSearchParams({
      key: defaultOneboundApiKey,
      secret: defaultOneboundApiSecret,
      q: keyword,
      start_price: String(startPrice),
      end_price: String(endPrice),
      page: String(page),
      page_size: String(pageSize),
      cat: "0",
      lang: "zh-CN"
    });

    const response = await fetch(`${defaultOneboundApiBase}${sourceConfig.path}?${params.toString()}`, {
      headers: {
        accept: "application/json, text/plain, */*",
        referer: "https://open.onebound.cn/"
      }
    });
    const payload = await response.json();

    if (!response.ok) {
      res.status(response.status).json({
        error: payload?.error || payload?.msg || payload?.message || "外部平台搜索失败"
      });
      return;
    }

    const normalized = normalizeExternalSearchPayload(payload, sourceConfig.label);
    const filteredItems = filterExternalItemsByMatchMode({
      items: normalized.items,
      keyword,
      matchMode
    });
    res.json({
      ok: true,
      source,
      matchMode,
      platformLabel: sourceConfig.label,
      keyword,
      page,
      pageSize,
      total: filteredItems.length,
      rawTotal: normalized.total,
      items: filteredItems
    });
  } catch (error) {
    res.status(500).json({ error: error.message || "外部平台搜索失败" });
  }
});

app.post("/api/public/upload/:token", (_req, res) => {
  res.status(410).json({ error: "旧的上传链接模式已停用，请在软件首页创建批次时直接上传多个货盘文件。" });
});

app.use((error, _req, res, _next) => {
  if (error instanceof multer.MulterError) {
    if (error.code === "LIMIT_FILE_SIZE") {
      res.status(400).json({ error: "文件太大了，请上传 50MB 以内的 Excel 或 CSV 文件" });
      return;
    }

    res.status(400).json({ error: error.message });
    return;
  }

  res.status(500).json({ error: error.message || "服务异常" });
});

app.listen(port, () => {
  console.log(`Supplier pricing hub is running at http://localhost:${port}`);
});

function initDb() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS collections (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      description TEXT,
      upload_token TEXT NOT NULL UNIQUE,
      created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS supplier_uploads (
      id TEXT PRIMARY KEY,
      collection_id TEXT NOT NULL,
      supplier_name TEXT NOT NULL,
      original_filename TEXT NOT NULL,
      stored_filename TEXT NOT NULL,
      row_count INTEGER NOT NULL DEFAULT 0,
      parsed_count INTEGER NOT NULL DEFAULT 0,
      category_summary TEXT,
      product_names_summary TEXT,
      catalog_overview TEXT,
      analysis_source TEXT,
      uploaded_at TEXT NOT NULL,
      FOREIGN KEY (collection_id) REFERENCES collections (id)
    );

    CREATE TABLE IF NOT EXISTS products (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      collection_id TEXT NOT NULL,
      upload_id TEXT NOT NULL,
      supplier_name TEXT NOT NULL,
      product_name TEXT,
      sku TEXT,
      spec TEXT,
      unit TEXT,
      price REAL,
      moq REAL,
      stock REAL,
      uploaded_at TEXT NOT NULL,
      raw_json TEXT,
      FOREIGN KEY (collection_id) REFERENCES collections (id),
      FOREIGN KEY (upload_id) REFERENCES supplier_uploads (id)
    );

    CREATE INDEX IF NOT EXISTS idx_products_collection ON products (collection_id);
    CREATE INDEX IF NOT EXISTS idx_products_sku ON products (sku);
    CREATE INDEX IF NOT EXISTS idx_products_name ON products (product_name);
    CREATE INDEX IF NOT EXISTS idx_uploads_collection ON supplier_uploads (collection_id);

    CREATE TABLE IF NOT EXISTS app_settings (
      key TEXT PRIMARY KEY,
      value TEXT
    );

    CREATE TABLE IF NOT EXISTS qjl_accounts (
      uid TEXT PRIMARY KEY,
      nickname TEXT,
      gh_id TEXT,
      gh_name TEXT,
      fans_num INTEGER NOT NULL DEFAULT 0,
      order_num INTEGER NOT NULL DEFAULT 0,
      mini_token TEXT NOT NULL,
      token_hint TEXT,
      user_info_json TEXT,
      homepage_list_json TEXT,
      feed_summary_json TEXT,
      profile_json TEXT,
      profile_summary TEXT,
      last_login_at TEXT NOT NULL,
      profile_updated_at TEXT,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );
  `);

  ensureTableColumn("supplier_uploads", "category_summary", "TEXT");
  ensureTableColumn("supplier_uploads", "product_names_summary", "TEXT");
  ensureTableColumn("supplier_uploads", "catalog_overview", "TEXT");
  ensureTableColumn("supplier_uploads", "analysis_source", "TEXT");
  ensureTableColumn("qjl_accounts", "nickname", "TEXT");
  ensureTableColumn("qjl_accounts", "gh_id", "TEXT");
  ensureTableColumn("qjl_accounts", "gh_name", "TEXT");
  ensureTableColumn("qjl_accounts", "fans_num", "INTEGER NOT NULL DEFAULT 0");
  ensureTableColumn("qjl_accounts", "order_num", "INTEGER NOT NULL DEFAULT 0");
  ensureTableColumn("qjl_accounts", "mini_token", "TEXT NOT NULL DEFAULT ''");
  ensureTableColumn("qjl_accounts", "token_hint", "TEXT");
  ensureTableColumn("qjl_accounts", "user_info_json", "TEXT");
  ensureTableColumn("qjl_accounts", "homepage_list_json", "TEXT");
  ensureTableColumn("qjl_accounts", "feed_summary_json", "TEXT");
  ensureTableColumn("qjl_accounts", "profile_json", "TEXT");
  ensureTableColumn("qjl_accounts", "profile_summary", "TEXT");
  ensureTableColumn("qjl_accounts", "last_login_at", "TEXT NOT NULL DEFAULT ''");
  ensureTableColumn("qjl_accounts", "profile_updated_at", "TEXT");
  ensureTableColumn("qjl_accounts", "created_at", "TEXT NOT NULL DEFAULT ''");
  ensureTableColumn("qjl_accounts", "updated_at", "TEXT NOT NULL DEFAULT ''");
}

function loadAiConfig() {
  const storedProvider = getSetting("llm_provider");
  const storedModel = getSetting("llm_model");
  const storedBaseUrl = getSetting("llm_base_url");
  const storedApiKey = getSetting("llm_api_key");

  aiRuntimeConfig.provider = storedProvider || defaultAiProvider;
  aiRuntimeConfig.model = storedModel || defaultAiModel;
  aiRuntimeConfig.baseUrl = (storedBaseUrl || defaultAiBaseUrl).replace(/\/$/, "");
  aiRuntimeConfig.apiKey = storedApiKey || process.env.DASHSCOPE_API_KEY || "";
  aiRuntimeConfig.source = storedApiKey ? "saved" : process.env.DASHSCOPE_API_KEY ? "environment" : "missing";
}

function getSetting(key) {
  const row = db.prepare("SELECT value FROM app_settings WHERE key = ?").get(key);
  return row?.value || "";
}

function persistSetting(key, value) {
  db.prepare(
    `
    INSERT INTO app_settings (key, value)
    VALUES (?, ?)
    ON CONFLICT(key) DO UPDATE SET value = excluded.value
    `
  ).run(key, value);
}

function deleteSetting(key) {
  db.prepare("DELETE FROM app_settings WHERE key = ?").run(key);
}

function getCurrentQjlUid() {
  return getSetting("current_qjl_uid");
}

function persistCurrentQjlUid(uid) {
  if (!uid) {
    deleteSetting("current_qjl_uid");
    return;
  }

  persistSetting("current_qjl_uid", uid);
}

function getPendingQjlLogin() {
  const key = getSetting("qjl_pending_login_key");
  const loginUrl = getSetting("qjl_pending_login_url");
  return {
    key,
    loginUrl,
    pending: Boolean(key && loginUrl)
  };
}

function persistPendingQjlLogin({ key, loginUrl }) {
  persistSetting("qjl_pending_login_key", key || "");
  persistSetting("qjl_pending_login_url", loginUrl || "");
}

function clearPendingQjlLogin() {
  deleteSetting("qjl_pending_login_key");
  deleteSetting("qjl_pending_login_url");
}

function getQjlAccountByUid(uid) {
  if (!uid) {
    return null;
  }

  return (
    db
      .prepare(
        `
        SELECT
          uid,
          nickname,
          gh_id,
          gh_name,
          fans_num,
          order_num,
          mini_token,
          token_hint,
          user_info_json,
          homepage_list_json,
          feed_summary_json,
          profile_json,
          profile_summary,
          last_login_at,
          profile_updated_at,
          created_at,
          updated_at
        FROM qjl_accounts
        WHERE uid = ?
        `
      )
      .get(uid) || null
  );
}

function getCurrentQjlAccount() {
  const currentUid = getCurrentQjlUid();
  if (currentUid) {
    const current = getQjlAccountByUid(currentUid);
    if (current) {
      return current;
    }
  }

  const latest =
    db
      .prepare(
        `
        SELECT
          uid,
          nickname,
          gh_id,
          gh_name,
          fans_num,
          order_num,
          mini_token,
          token_hint,
          user_info_json,
          homepage_list_json,
          feed_summary_json,
          profile_json,
          profile_summary,
          last_login_at,
          profile_updated_at,
          created_at,
          updated_at
        FROM qjl_accounts
        ORDER BY updated_at DESC, last_login_at DESC
        LIMIT 1
        `
      )
      .get() || null;

  if (latest?.uid) {
    persistCurrentQjlUid(latest.uid);
  }

  return latest;
}

function persistQjlAccount(account) {
  db.prepare(
    `
    INSERT INTO qjl_accounts (
      uid,
      nickname,
      gh_id,
      gh_name,
      fans_num,
      order_num,
      mini_token,
      token_hint,
      user_info_json,
      homepage_list_json,
      feed_summary_json,
      profile_json,
      profile_summary,
      last_login_at,
      profile_updated_at,
      created_at,
      updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(uid) DO UPDATE SET
      nickname = excluded.nickname,
      gh_id = excluded.gh_id,
      gh_name = excluded.gh_name,
      fans_num = excluded.fans_num,
      order_num = excluded.order_num,
      mini_token = excluded.mini_token,
      token_hint = excluded.token_hint,
      user_info_json = excluded.user_info_json,
      homepage_list_json = excluded.homepage_list_json,
      feed_summary_json = excluded.feed_summary_json,
      profile_json = excluded.profile_json,
      profile_summary = excluded.profile_summary,
      last_login_at = excluded.last_login_at,
      profile_updated_at = excluded.profile_updated_at,
      updated_at = excluded.updated_at
    `
  ).run(
    account.uid,
    account.nickname || "",
    account.ghId || "",
    account.ghName || "",
    Number(account.fansNum || 0),
    Number(account.orderNum || 0),
    account.miniToken || "",
    account.tokenHint || maskSecret(account.miniToken || ""),
    account.userInfoJson || "",
    account.homepageListJson || "",
    account.feedSummaryJson || "",
    account.profileJson || "",
    account.profileSummary || "",
    account.lastLoginAt,
    account.profileUpdatedAt || "",
    account.createdAt,
    account.updatedAt
  );

  persistCurrentQjlUid(account.uid);
}

function deleteCurrentQjlAccount() {
  const current = getCurrentQjlAccount();
  if (current?.uid) {
    db.prepare("DELETE FROM qjl_accounts WHERE uid = ?").run(current.uid);
  }

  persistCurrentQjlUid("");
  clearPendingQjlLogin();
}

function formatQjlAccount(row) {
  if (!row) {
    return null;
  }

  const profile = safeParseJson(row.profile_json) || {};
  const feedSummary = safeParseJson(row.feed_summary_json) || {};
  const homepageList = normalizeQjlHomes(safeParseJson(row.homepage_list_json) || {});
  const currentHomepage = homepageList.find((item) => String(item?.ghCode || "") === String(row.gh_id || "")) || null;

  return {
    uid: String(row.uid || ""),
    nickname: String(row.nickname || ""),
    ghId: String(row.gh_id || ""),
    ghName: String(row.gh_name || ""),
    fansNum: Number(row.fans_num || 0),
    orderNum: Number(row.order_num || 0),
    tokenHint: String(row.token_hint || ""),
    lastLoginAt: row.last_login_at || "",
    profileUpdatedAt: row.profile_updated_at || "",
    profileSummary: String(row.profile_summary || ""),
    homepageList: homepageList.map(formatQjlHomepageOption),
    currentHomepage: currentHomepage ? formatQjlHomepageOption(currentHomepage) : null,
    profile,
    feedSummary
  };
}

function formatQjlHomepageOption(home) {
  return {
    ghCode: String(home?.ghCode || "").trim(),
    ghName: String(home?.ghName || "").trim(),
    fansNum: toSafeInt(home?.fansNum),
    orderNum: toSafeInt(home?.ghOrderNum),
    ghType: toSafeInt(home?.ghType),
    useStatus: toSafeInt(home?.useStatus)
  };
}

function ensureTableColumn(tableName, columnName, columnDefinition) {
  const columns = db.prepare(`PRAGMA table_info(${tableName})`).all();
  if (columns.some((column) => column.name === columnName)) {
    return;
  }

  db.exec(`ALTER TABLE ${tableName} ADD COLUMN ${columnName} ${columnDefinition}`);
}

function ensureDir(directory) {
  fs.mkdirSync(directory, { recursive: true });
}

function cleanupUpload(file) {
  if (!file?.path) {
    return;
  }

  if (fs.existsSync(file.path)) {
    fs.unlinkSync(file.path);
  }
}

function cleanupUploads(files) {
  for (const file of files || []) {
    cleanupUpload(file);
  }
}

function cleanupStoredUploadsByCollection(collectionId) {
  if (!collectionId) {
    return;
  }

  const directory = path.join(catalogUploadDir, collectionId);
  if (!fs.existsSync(directory)) {
    return;
  }

  fs.rmSync(directory, { recursive: true, force: true });
}

function moveUploadIntoCollectionFolder({ collectionId, file }) {
  const collectionDir = path.join(catalogUploadDir, collectionId);
  ensureDir(collectionDir);

  const nextPath = path.join(collectionDir, file.filename);
  fs.renameSync(file.path, nextPath);

  return {
    absolutePath: nextPath,
    relativePath: path.join(collectionId, file.filename)
  };
}

function resolveStoredUploadPath(upload) {
  const stored = String(upload?.stored_filename || "").trim();
  if (!stored) {
    return "";
  }

  const candidatePaths = [];
  if (stored.includes("/") || stored.includes("\\")) {
    candidatePaths.push(path.join(catalogUploadDir, stored));
  } else {
    candidatePaths.push(path.join(catalogUploadDir, String(upload.collection_id || ""), stored));
    candidatePaths.push(path.join(tempUploadDir, stored));
    candidatePaths.push(path.join(legacyUploadDir, stored));
  }

  return candidatePaths.find((filePath) => fs.existsSync(filePath)) || candidatePaths[0] || "";
}

function supplierNameFromFilename(originalname) {
  const filename = path.basename(String(originalname || ""), path.extname(String(originalname || "")));
  return filename.trim().replace(/[_-]+/g, " ").replace(/\s+/g, " ") || "未命名供应商";
}

function importUploadedFile({ collection, supplierName, file }) {
  const parsed = parseWorkbook(file.path, supplierName);

  if (parsed.parsedCount === 0) {
    throw new Error("没有识别到可导入的商品数据，请检查表头是否包含商品名和价格等字段");
  }

  const uploadId = crypto.randomUUID();
  const now = new Date().toISOString();

  const insertUpload = db.prepare(
    `
    INSERT INTO supplier_uploads (
      id,
      collection_id,
      supplier_name,
      original_filename,
      stored_filename,
      row_count,
      parsed_count,
      uploaded_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `
  );

  const insertProduct = db.prepare(
    `
    INSERT INTO products (
      collection_id,
      upload_id,
      supplier_name,
      product_name,
      sku,
      spec,
      unit,
      price,
      moq,
      stock,
      uploaded_at,
      raw_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `
  );

  const transaction = db.transaction(() => {
    insertUpload.run(
      uploadId,
      collection.id,
      supplierName,
      file.originalname,
      file.filename,
      parsed.rowCount,
      parsed.parsedCount,
      now
    );

    for (const record of parsed.items) {
      insertProduct.run(
        collection.id,
        uploadId,
        supplierName,
        record.productName,
        record.sku,
        record.spec,
        record.unit,
        record.price,
        record.moq,
        record.stock,
        now,
        JSON.stringify(record.raw)
      );
    }
  });

  transaction();
  return parsed;
}

function importBatchFilesIntoCollection({ collection, parsedFiles }) {
  const insertCollection = db.prepare(
    `
    INSERT INTO collections (id, name, description, upload_token, created_at)
    VALUES (?, ?, ?, ?, ?)
    `
  );

  const insertUpload = db.prepare(
    `
    INSERT INTO supplier_uploads (
      id,
      collection_id,
      supplier_name,
      original_filename,
      stored_filename,
      row_count,
      parsed_count,
      category_summary,
      product_names_summary,
      catalog_overview,
      analysis_source,
      uploaded_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `
  );

  const insertProduct = db.prepare(
    `
    INSERT INTO products (
      collection_id,
      upload_id,
      supplier_name,
      product_name,
      sku,
      spec,
      unit,
      price,
      moq,
      stock,
      uploaded_at,
      raw_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `
  );

  let totalRowCount = 0;
  let totalParsedCount = 0;

  const transaction = db.transaction(() => {
    insertCollection.run(
      collection.id,
      collection.name,
      collection.description,
      collection.upload_token,
      collection.created_at
    );

    for (const item of parsedFiles) {
      const uploadId = crypto.randomUUID();
      const uploadedAt = new Date().toISOString();
      totalRowCount += item.parsed.rowCount;
      totalParsedCount += item.parsed.parsedCount;

      insertUpload.run(
        uploadId,
        collection.id,
        item.supplierName,
        item.file.originalname,
        item.storedFile.relativePath,
        item.parsed.rowCount,
        item.parsed.parsedCount,
        item.summary.categorySummary,
        item.summary.productNamesSummary,
        item.summary.catalogOverview,
        item.summary.analysisSource,
        uploadedAt
      );

      for (const record of item.parsed.items) {
        insertProduct.run(
          collection.id,
          uploadId,
          item.supplierName,
          record.productName,
          record.sku,
          record.spec,
          record.unit,
          record.price,
          record.moq,
          record.stock,
          uploadedAt,
          JSON.stringify(record.raw)
        );
      }
    }
  });

  transaction();
  return { totalRowCount, totalParsedCount };
}

function safeParseJson(value) {
  if (!value) {
    return null;
  }

  try {
    return JSON.parse(value);
  } catch (_error) {
    return null;
  }
}

function compactRawRow(raw) {
  const source = raw && typeof raw === "object" ? raw : {};
  const compact = {};

  for (const [key, value] of Object.entries(source).slice(0, 12)) {
    const normalizedKey = String(key || "").trim();
    if (!normalizedKey) {
      continue;
    }

    const normalizedValue = stringifyAiCell(value).slice(0, 80);
    if (!normalizedValue) {
      continue;
    }

    compact[normalizedKey] = normalizedValue;
  }

  return compact;
}

function maskSecret(value) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }

  if (text.length <= 8) {
    return "已保存";
  }

  return `${text.slice(0, 4)}...${text.slice(-4)}`;
}

async function requestJson(url, options = {}) {
  const response = await fetch(url, options);
  const text = await response.text();

  if (!response.ok) {
    throw new Error(`请求失败：${response.status} ${text}`);
  }

  try {
    return JSON.parse(text);
  } catch (_error) {
    throw new Error(`接口没有返回合法 JSON：${text.slice(0, 200)}`);
  }
}

function buildQjlHeaders({ token = "", uid = "", miniRoute = qjlMiniRoutes.homepage } = {}) {
  const headers = {
    ...qjlBaseHeaders,
    "mini-route": miniRoute
  };

  if (token) {
    headers.Authorization = token;
  }

  if (uid) {
    headers.uid = String(uid);
  }

  return headers;
}

function pickQjlLoginKey(payload) {
  return String(payload?.key || payload?.id || payload?.data?.key || "").trim();
}

function pickQjlMiniToken(payload) {
  return String(payload?.miniToken || payload?.data?.miniToken || "").trim();
}

function pickQjlUid(payload) {
  return String(payload?.uid || payload?.data?.uid || "").trim();
}

async function requestQjlLoginKey() {
  const payload = await requestJson(`${defaultQjlAuthBase.replace(/\/$/, "")}/get_login_key`);
  const key = pickQjlLoginKey(payload);

  if (!key) {
    throw new Error("群接龙登录服务没有返回有效 key");
  }

  return {
    key,
    loginUrl: `${defaultQjlAuthBase.replace(/\/$/, "")}/login?key=${encodeURIComponent(key)}`
  };
}

async function pollQjlLoginByKey(key) {
  const payload = await requestJson(
    `${defaultQjlAuthBase.replace(/\/$/, "")}/getTokenByKey?key=${encodeURIComponent(key)}`
  );

  return {
    key,
    uid: pickQjlUid(payload),
    miniToken: pickQjlMiniToken(payload),
    raw: payload
  };
}

async function callQjlApi({
  path: apiPath,
  method = "GET",
  token = "",
  uid = "",
  miniRoute = qjlMiniRoutes.homepage,
  body
}) {
  const url = apiPath.startsWith("http") ? apiPath : `${defaultQjlApiBase}${apiPath}`;
  const options = {
    method,
    headers: buildQjlHeaders({ token, uid, miniRoute })
  };

  if (body !== undefined) {
    options.body = JSON.stringify(body);
  }

  return requestJson(url, options);
}

function unwrapQjlData(payload) {
  if (payload && typeof payload === "object" && payload.data && typeof payload.data === "object") {
    return payload.data;
  }

  return payload && typeof payload === "object" ? payload : {};
}

function normalizeQjlHomes(data) {
  const merged = [];
  const seen = new Set();

  if (Array.isArray(data)) {
    for (const home of data) {
      const ghCode = String(home?.ghCode || "").trim();
      if (!ghCode || seen.has(ghCode)) {
        continue;
      }
      seen.add(ghCode);
      merged.push(home);
    }
    return merged;
  }

  for (const key of ["allHomeList", "normalHomeList", "expireHomeList"]) {
    const homes = data?.[key];
    if (!Array.isArray(homes)) {
      continue;
    }

    for (const home of homes) {
      const ghCode = String(home?.ghCode || "").trim();
      if (!ghCode || seen.has(ghCode)) {
        continue;
      }

      seen.add(ghCode);
      merged.push(home);
    }
  }

  return merged;
}

function loadLocalEnvFiles(filePaths) {
  for (const filePath of filePaths) {
    if (!filePath || !fs.existsSync(filePath)) {
      continue;
    }

    const lines = fs.readFileSync(filePath, "utf8").split(/\r?\n/);
    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#")) {
        continue;
      }

      const match = trimmed.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
      if (!match) {
        continue;
      }

      const [, key, rawValue] = match;
      if (process.env[key]) {
        continue;
      }

      let value = rawValue.trim();
      if (
        (value.startsWith("\"") && value.endsWith("\"")) ||
        (value.startsWith("'") && value.endsWith("'"))
      ) {
        value = value.slice(1, -1);
      }

      process.env[key] = value;
    }
  }
}

function toSafeInt(value) {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.trunc(value);
  }

  const digits = String(value).match(/-?\d+/);
  return digits ? Number(digits[0]) : 0;
}

function selectBestQjlHome(homes) {
  if (!Array.isArray(homes) || !homes.length) {
    return null;
  }

  return [...homes].sort((a, b) => {
    const scoreA = [
      toSafeInt(a.useStatus) === 10 ? 1 : 0,
      toSafeInt(a.ghType) === 20 ? 1 : 0,
      toSafeInt(a.fansNum),
      toSafeInt(a.ghOrderNum)
    ];
    const scoreB = [
      toSafeInt(b.useStatus) === 10 ? 1 : 0,
      toSafeInt(b.ghType) === 20 ? 1 : 0,
      toSafeInt(b.fansNum),
      toSafeInt(b.ghOrderNum)
    ];

    for (let i = 0; i < scoreA.length; i += 1) {
      if (scoreB[i] !== scoreA[i]) {
        return scoreB[i] - scoreA[i];
      }
    }
    return 0;
  })[0];
}

async function fetchQjlHomepageList({ token, uid }) {
  const payload = await callQjlApi({
    path: "/ghome-major/switch/get_switch_home_list_v2",
    method: "GET",
    token,
    uid,
    miniRoute: qjlMiniRoutes.homepageList
  });
  const data = unwrapQjlData(payload);
  return normalizeQjlHomes(data);
}

async function fetchQjlFeedPage({ token, uid, ghId, page, pageSize }) {
  const payload = await callQjlApi({
    path: "/ghome-feed/ghome_feed/query_ghome_feed_back_v2",
    method: "POST",
    token,
    uid,
    miniRoute: qjlMiniRoutes.homepage,
    body: {
      page,
      pageSize,
      ghId,
      keyword: "",
      ghType: 20,
      sortRule: 10
    }
  });

  const data = unwrapQjlData(payload);
  const feedData = data?.data && typeof data.data === "object" ? data.data : data;

  return {
    items: Array.isArray(feedData?.feedItemDTOList) ? feedData.feedItemDTOList : [],
    hasNextPage: Boolean(feedData?.hasNextPage)
  };
}

async function fetchQjlFeed({ token, uid, ghId, pageSize = 90 }) {
  const merged = [];
  const chunkSize = Math.min(pageSize, 30);
  let page = 1;

  while (merged.length < pageSize) {
    const next = await fetchQjlFeedPage({
      token,
      uid,
      ghId,
      page,
      pageSize: chunkSize
    });

    if (!next.items.length) {
      break;
    }

    merged.push(...next.items);
    if (!next.hasNextPage) {
      break;
    }
    page += 1;
  }

  return merged.slice(0, pageSize);
}

async function fetchQjlBusinessOverview({ token, uid, ghId }) {
  const payload = await callQjlApi({
    path: "/ghdata-biz/api/ghData/query",
    method: "POST",
    token,
    uid,
    miniRoute: qjlMiniRoutes.dataOverview,
    body: {
      module: "trade",
      ghId,
      ghViewType: 20,
      resultMode: "point",
      dataNames: ["orderNum", "realPrice", "visitUserNum", "buyerNum", "conversionRate"],
      dateTag: "last30"
    }
  });

  const data = unwrapQjlData(payload);
  return data?.data && typeof data.data === "object" ? data.data : data;
}

function inferQjlCategory(text) {
  const lowered = String(text || "").toLowerCase();
  for (const [category, keywords] of Object.entries(qjlProfileCategoryKeywords)) {
    if (keywords.some((keyword) => lowered.includes(String(keyword).toLowerCase()))) {
      return category;
    }
  }
  return "其他";
}

function stripTimePrefix(text) {
  return String(text || "").replace(/^\d{1,2}\/\d{1,2}\s+\d{1,2}:\d{1,2}\S*\s*/, "").trim();
}

function extractQjlProductName(title) {
  const cleaned = stripTimePrefix(title);
  for (const [left, right] of [["《", "》"], ["【", "】"], ["「", "」"]]) {
    const match = cleaned.match(new RegExp(`${left}([^${left}${right}]+)${right}`));
    if (match?.[1]) {
      return match[1].trim();
    }
  }

  return cleaned.split(/[|｜/]/)[0].trim().slice(0, 60);
}

function extractQjlStats(item, act) {
  const sources = [
    act?.feedGroupSaleChildActDTO,
    act?.feedGroupGrandSonActDTO,
    act?.feedGroupActDTO,
    act?.feedApplyActDTO,
    act?.feedChartActDTO,
    item?.feedGroupSaleChildActDTO,
    item?.feedGroupGrandSonActDTO,
    item?.feedGroupActDTO,
    item?.feedApplyActDTO,
    item?.feedChartActDTO
  ];

  for (const source of sources) {
    if (!source || typeof source !== "object") {
      continue;
    }

    if ("orderCount" in source || "viewCount" in source) {
      return {
        orderCount: toSafeInt(source.orderCount ?? act?.participationCount),
        viewCount: toSafeInt(source.viewCount ?? act?.viewCount)
      };
    }
  }

  return {
    orderCount: toSafeInt(act?.participationCount),
    viewCount: toSafeInt(act?.viewCount)
  };
}

function extractQjlSourceLeader(item, act) {
  const candidates = [
    act?.nickName,
    act?.ghName,
    act?.feedGroupActDTO?.groupSalerRootActInfoDTO?.ghName,
    act?.feedGroupSaleChildActDTO?.groupSalerRootActInfoDTO?.ghName,
    item?.feedGroupActDTO?.groupSalerRootActInfoDTO?.ghName,
    item?.feedGroupSaleChildActDTO?.groupSalerRootActInfoDTO?.ghName
  ];

  return candidates.find((value) => String(value || "").trim()) || "";
}

function normalizeQjlActivity(item) {
  const act = item?.actItemDTO || {};
  const actId = String(act.actId || "").trim();
  if (!actId) {
    return null;
  }

  const title = String(act.activityName || "").trim();
  const { orderCount, viewCount } = extractQjlStats(item, act);
  const productName = extractQjlProductName(title);

  return {
    actId,
    title,
    productName,
    category: inferQjlCategory(productName || title),
    sourceLeader: extractQjlSourceLeader(item, act),
    orderCount,
    viewCount,
    createTime: toSafeInt(act.createTime)
  };
}

function summarizeQjlFeedActivities(activities) {
  const normalized = activities.map(normalizeQjlActivity).filter(Boolean);
  const categoryCounter = new Map();
  const leaderCounter = new Map();
  const recentCutoff = Date.now() - 365 * 24 * 60 * 60 * 1000;

  for (const item of normalized) {
    categoryCounter.set(item.category, (categoryCounter.get(item.category) || 0) + 1);
    if (item.sourceLeader) {
      leaderCounter.set(item.sourceLeader, (leaderCounter.get(item.sourceLeader) || 0) + 1);
    }
  }

  const topCategories = [...categoryCounter.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 6)
    .map(([name, count]) => ({ name, count }));

  const topLeaders = [...leaderCounter.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([name, count]) => ({ name, count }));

  const bestSelling = [...normalized]
    .sort((a, b) => (b.orderCount - a.orderCount) || (b.viewCount - a.viewCount))
    .slice(0, 5)
    .map((item) => ({
      title: item.title,
      productName: item.productName,
      orderCount: item.orderCount,
      viewCount: item.viewCount
    }));

  const mostViewed = [...normalized]
    .sort((a, b) => (b.viewCount - a.viewCount) || (b.orderCount - a.orderCount))
    .slice(0, 5)
    .map((item) => ({
      title: item.title,
      productName: item.productName,
      orderCount: item.orderCount,
      viewCount: item.viewCount
    }));

  return {
    activityCount: normalized.length,
    recentCount: normalized.filter((item) => item.createTime >= recentCutoff).length,
    totalOrders: normalized.reduce((sum, item) => sum + item.orderCount, 0),
    totalViews: normalized.reduce((sum, item) => sum + item.viewCount, 0),
    topCategories,
    topLeaders,
    bestSelling,
    mostViewed,
    sampleActivities: normalized.slice(0, 12).map((item) => ({
      title: item.title,
      category: item.category,
      sourceLeader: item.sourceLeader,
      orderCount: item.orderCount,
      viewCount: item.viewCount
    }))
  };
}

function normalizeStringList(values, limit = 6) {
  return uniqueStrings((Array.isArray(values) ? values : []).map(stringifyAiCell)).slice(0, limit);
}

function buildQjlPortraitFallback({ uid, ghName, home, feedSummary, overview }) {
  const topCategories = (feedSummary.topCategories || []).map((item) => item.name).slice(0, 4);
  const summary = [
    ghName ? `当前登录主页为 ${ghName}` : `当前已登录群接龙 uid ${uid}`,
    topCategories.length ? `近期偏向 ${topCategories.join("、")}` : "近期经营方向仍待补充更多样本",
    feedSummary.activityCount ? `主页样本活动 ${feedSummary.activityCount} 条` : "暂无足够活动样本",
    toSafeInt(home?.fansNum) ? `粉丝约 ${toSafeInt(home?.fansNum)}` : ""
  ]
    .filter(Boolean)
    .join("，");

  return {
    summary,
    tags: normalizeStringList([
      topCategories[0] || "",
      toSafeInt(home?.ghType) === 20 ? "帮卖主页" : "个人主页",
      feedSummary.recentCount ? "近一年有持续活动" : "近期活动偏少"
    ]),
    preferredCategories: topCategories,
    businessStage:
      feedSummary.activityCount >= 20 ? "已有一定经营样本" : feedSummary.activityCount ? "轻量经营中" : "待补充经营样本",
    sourcingPreferences: normalizeStringList((feedSummary.topLeaders || []).map((item) => item.name)),
    analysisFocus: normalizeStringList([
      "优先看和主页偏好品类匹配的货盘",
      "关注价格带是否适合当前主页用户",
      "优先识别缺货、缺价和补采风险"
    ]),
    cautionPoints: normalizeStringList([
      !feedSummary.activityCount ? "群接龙活动样本不足，画像可信度有限" : "",
      !overview || !Object.keys(overview).length ? "未拿到完整经营概览，建议后续补充" : ""
    ]),
    source: "rule"
  };
}

async function buildQjlPortrait({ uid, ghId, ghName, home, feedSummary, overview }) {
  const fallback = buildQjlPortraitFallback({ uid, ghName, home, feedSummary, overview });

  if (!aiRuntimeConfig.apiKey) {
    return fallback;
  }

  try {
    const result = await callDashScopeJson({
      systemPrompt: [
        "你是一个群接龙团长画像助手。",
        "请根据登录用户的群接龙主页信息、近期活动和经营概览，生成一份适合货盘分析场景使用的用户画像。",
        "必须输出 JSON，不要输出任何额外说明。",
        "返回结构必须是 { summary, tags, preferred_categories, business_stage, sourcing_preferences, analysis_focus, caution_points }。",
        "summary 为一段不超过 90 字的中文总结，其他字段均为字符串数组或短字符串。"
      ].join(" "),
      userPrompt: JSON.stringify(
        {
          uid,
          ghId,
          ghName,
          home: {
            ghType: toSafeInt(home?.ghType),
            fansNum: toSafeInt(home?.fansNum),
            ghOrderNum: toSafeInt(home?.ghOrderNum),
            useStatus: toSafeInt(home?.useStatus)
          },
          feedSummary,
          overview
        },
        null,
        2
      )
    });

    return {
      summary: stringifyAiCell(result.summary) || fallback.summary,
      tags: normalizeStringList(result.tags),
      preferredCategories: normalizeStringList(result.preferred_categories),
      businessStage: stringifyAiCell(result.business_stage) || fallback.businessStage,
      sourcingPreferences: normalizeStringList(result.sourcing_preferences),
      analysisFocus: normalizeStringList(result.analysis_focus),
      cautionPoints: normalizeStringList(result.caution_points),
      source: "ai"
    };
  } catch (_error) {
    return fallback;
  }
}

async function buildQjlAccountFromLogin({ uid, miniToken }) {
  const now = new Date().toISOString();
  let homes = [];
  let selectedHome = null;
  let ghId = "";
  let ghName = "";
  let overview = {};
  let feedSummary = {
    activityCount: 0,
    recentCount: 0,
    totalOrders: 0,
    totalViews: 0,
    topCategories: [],
    topLeaders: [],
    bestSelling: [],
    mostViewed: [],
    sampleActivities: []
  };

  try {
    homes = await fetchQjlHomepageList({ token: miniToken, uid });
    selectedHome = selectBestQjlHome(homes);
    ghId = String(selectedHome?.ghCode || "").trim();
    ghName = String(selectedHome?.ghName || "").trim();
  } catch (_error) {
    homes = [];
  }

  if (ghId) {
    try {
      const activities = await fetchQjlFeed({ token: miniToken, uid, ghId, pageSize: 100 });
      feedSummary = summarizeQjlFeedActivities(activities);
    } catch (_error) {
      feedSummary = {
        ...feedSummary
      };
    }

    try {
      overview = await fetchQjlBusinessOverview({ token: miniToken, uid, ghId });
    } catch (_error) {
      overview = {};
    }
  }

  const profile = await buildQjlPortrait({
    uid,
    ghId,
    ghName: ghName || `群接龙用户 ${uid}`,
    home: selectedHome,
    feedSummary,
    overview
  });

  return {
    uid: String(uid),
    nickname: ghName || `群接龙用户 ${uid}`,
    ghId,
    ghName: ghName || `群接龙用户 ${uid}`,
    fansNum: toSafeInt(selectedHome?.fansNum),
    orderNum: toSafeInt(selectedHome?.ghOrderNum),
    miniToken,
    tokenHint: maskSecret(miniToken),
    userInfoJson: JSON.stringify({
      uid: String(uid),
      ghId,
      ghName: ghName || "",
      loginSource: "qjlMiniSkills-compatible"
    }),
    homepageListJson: JSON.stringify(homes),
    feedSummaryJson: JSON.stringify({
      ...feedSummary,
      overview
    }),
    profileJson: JSON.stringify(profile),
    profileSummary: profile.summary || "",
    lastLoginAt: now,
    profileUpdatedAt: now,
    createdAt: now,
    updatedAt: now
  };
}

async function rebuildQjlAccountForHomepage({ account, ghCode }) {
  const uid = String(account?.uid || "").trim();
  const miniToken = String(account?.mini_token || "").trim();
  if (!uid || !miniToken) {
    throw new Error("当前群接龙登录状态无效，请重新登录。");
  }

  let homes = normalizeQjlHomes(safeParseJson(account.homepage_list_json) || {});
  if (!homes.length) {
    homes = await fetchQjlHomepageList({ token: miniToken, uid });
  }

  const selectedHome = homes.find((home) => String(home?.ghCode || "").trim() === ghCode);
  if (!selectedHome) {
    throw new Error("未找到对应的群接龙主页，请先刷新主页列表。");
  }

  const now = new Date().toISOString();
  const safeGhCode = String(selectedHome.ghCode || "").trim();
  const ghName = String(selectedHome.ghName || "").trim();

  let feedSummary = {
    activityCount: 0,
    recentCount: 0,
    totalOrders: 0,
    totalViews: 0,
    topCategories: [],
    topLeaders: [],
    bestSelling: [],
    mostViewed: [],
    sampleActivities: []
  };
  let overview = {};

  try {
    const activities = await fetchQjlFeed({ token: miniToken, uid, ghId: safeGhCode, pageSize: 100 });
    feedSummary = summarizeQjlFeedActivities(activities);
  } catch (_error) {
    feedSummary = { ...feedSummary };
  }

  try {
    overview = await fetchQjlBusinessOverview({ token: miniToken, uid, ghId: safeGhCode });
  } catch (_error) {
    overview = {};
  }

  const profile = await buildQjlPortrait({
    uid,
    ghId: safeGhCode,
    ghName: ghName || `群接龙用户 ${uid}`,
    home: selectedHome,
    feedSummary,
    overview
  });

  return {
    uid,
    nickname: ghName || String(account.nickname || "") || `群接龙用户 ${uid}`,
    ghId: safeGhCode,
    ghName: ghName || `群接龙用户 ${uid}`,
    fansNum: toSafeInt(selectedHome?.fansNum),
    orderNum: toSafeInt(selectedHome?.ghOrderNum),
    miniToken,
    tokenHint: String(account.token_hint || "") || maskSecret(miniToken),
    userInfoJson: JSON.stringify({
      uid,
      ghId: safeGhCode,
      ghName: ghName || "",
      loginSource: "qjlMiniSkills-compatible"
    }),
    homepageListJson: JSON.stringify(homes),
    feedSummaryJson: JSON.stringify({
      ...feedSummary,
      overview
    }),
    profileJson: JSON.stringify(profile),
    profileSummary: profile.summary || "",
    lastLoginAt: account.last_login_at || now,
    profileUpdatedAt: now,
    createdAt: account.created_at || now,
    updatedAt: now
  };
}

function getCurrentQjlProfileContext() {
  const account = getCurrentQjlAccount();
  if (!account?.uid || !account?.profile_json) {
    return null;
  }

  return {
    uid: String(account.uid || ""),
    ghId: String(account.gh_id || ""),
    ghName: String(account.gh_name || ""),
    fansNum: Number(account.fans_num || 0),
    orderNum: Number(account.order_num || 0),
    profileSummary: String(account.profile_summary || ""),
    profile: safeParseJson(account.profile_json) || {},
    feedSummary: safeParseJson(account.feed_summary_json) || {}
  };
}

function getCollectionById(id) {
  return db
    .prepare(
      `
      SELECT
        c.id,
        c.name,
        c.description,
        c.upload_token,
        c.created_at,
        COUNT(DISTINCT su.id) AS upload_count,
        COUNT(DISTINCT p.id) AS product_count
      FROM collections c
      LEFT JOIN supplier_uploads su ON su.collection_id = c.id
      LEFT JOIN products p ON p.collection_id = c.id
      WHERE c.id = ?
      GROUP BY c.id
      `
    )
    .get(id);
}

function formatCollection(row) {
  if (!row) {
    return null;
  }

  return {
    id: row.id,
    name: row.name,
    description: row.description,
    createdAt: row.created_at,
    uploadCount: Number(row.upload_count || 0),
    productCount: Number(row.product_count || 0),
    uploadToken: row.upload_token,
    uploadUrl: `/u/${row.upload_token}`
  };
}

function buildUploadBatchItem({ collection, upload, products }) {
  const inferredCategories = inferCategoryLabels(products);
  const categoryLabels = parseSummaryList(upload.category_summary) || inferredCategories;
  const productNames = parseSummaryList(upload.product_names_summary) || inferProductNames(products);

  return {
    id: upload.id,
    supplierName: upload.supplier_name,
    catalogName: upload.original_filename,
    storedPath: upload.stored_filename || "",
    fileUrl: `/api/uploads/${upload.id}/file`,
    batchName: collection.name,
    uploadedAt: upload.uploaded_at,
    productCount: Number(upload.parsed_count || products.length || 0),
    categoryCount: categoryLabels.length,
    categoryLabels,
    categorySummary: categoryLabels.length ? categoryLabels.join(" / ") : "未识别",
    productNames,
    productNamesSummary: productNames.length ? productNames.join(" / ") : "未识别",
    catalogOverview: String(upload.catalog_overview || "").trim() || `共导入 ${Number(upload.parsed_count || products.length || 0)} 个商品`,
    analysisSource: String(upload.analysis_source || "").trim() || "rule"
  };
}

function inferCategoryLabels(products) {
  const labels = [];
  const used = new Set();

  for (const row of products || []) {
    const raw = safeParseJson(row.raw_json);
    const values = extractByAliases(raw, headerAliases.category);

    for (const value of values) {
      const normalized = normalizeText(value);
      if (!normalized || used.has(normalized)) {
        continue;
      }

      used.add(normalized);
      labels.push(value);

      if (labels.length >= 6) {
        return labels;
      }
    }
  }

  return labels;
}

function inferProductNames(products) {
  const names = [];
  const used = new Set();

  for (const row of products || []) {
    const text = String(row.product_name || "").trim();
    const normalized = normalizeText(text);
    if (!normalized || used.has(normalized)) {
      continue;
    }

    used.add(normalized);
    names.push(text);

    if (names.length >= 6) {
      break;
    }
  }

  return names;
}

function parseSummaryList(value) {
  const text = String(value || "").trim();
  if (!text) {
    return [];
  }

  return text
    .split(/[\n,，/、]+/)
    .map((item) => String(item || "").trim())
    .filter(Boolean)
    .slice(0, 6);
}

function extractByAliases(row, aliases) {
  if (!row || typeof row !== "object") {
    return [];
  }

  const values = [];
  for (const [key, value] of Object.entries(row)) {
    const normalizedKey = normalizeHeader(key);
    if (!aliases.some((alias) => headerMatches(normalizedKey, alias))) {
      continue;
    }

    const text = String(value || "").trim();
    if (!text) {
      continue;
    }

    values.push(text);
  }

  return values;
}

function getUploadsByIds(collectionId, uploadIds) {
  if (!uploadIds.length) {
    return [];
  }

  const placeholders = uploadIds.map(() => "?").join(", ");
  return db
    .prepare(
      `
      SELECT
        id,
        supplier_name,
        original_filename,
        parsed_count,
        uploaded_at
      FROM supplier_uploads
      WHERE collection_id = ?
        AND id IN (${placeholders})
      ORDER BY uploaded_at DESC, id DESC
      `
    )
    .all(collectionId, ...uploadIds);
}

function getProductsByUploadIds(collectionId, uploadIds) {
  if (!uploadIds.length) {
    return [];
  }

  const placeholders = uploadIds.map(() => "?").join(", ");
  return db
    .prepare(
      `
      SELECT
        upload_id,
        supplier_name,
        product_name,
        sku,
        spec,
        unit,
        price,
        moq,
        stock,
        uploaded_at,
        raw_json
      FROM products
      WHERE collection_id = ?
        AND upload_id IN (${placeholders})
      ORDER BY uploaded_at DESC, supplier_name ASC, id ASC
      `
    )
    .all(collectionId, ...uploadIds);
}

function getUploadsByIdsAcrossCollections(uploadIds) {
  if (!uploadIds.length) {
    return [];
  }

  const placeholders = uploadIds.map(() => "?").join(", ");
  return db
    .prepare(
      `
      SELECT
        su.id,
        su.collection_id,
        su.supplier_name,
        su.original_filename,
        su.parsed_count,
        su.uploaded_at,
        c.name AS collection_name,
        c.description AS collection_description
      FROM supplier_uploads su
      INNER JOIN collections c ON c.id = su.collection_id
      WHERE su.id IN (${placeholders})
      ORDER BY su.uploaded_at DESC, su.id DESC
      `
    )
    .all(...uploadIds);
}

function getProductsByUploadIdsAcrossCollections(uploadIds) {
  if (!uploadIds.length) {
    return [];
  }

  const placeholders = uploadIds.map(() => "?").join(", ");
  return db
    .prepare(
      `
      SELECT
        upload_id,
        supplier_name,
        product_name,
        sku,
        spec,
        unit,
        price,
        moq,
        stock,
        uploaded_at,
        raw_json
      FROM products
      WHERE upload_id IN (${placeholders})
      ORDER BY uploaded_at DESC, supplier_name ASC, id ASC
      `
    )
    .all(...uploadIds);
}

function parseWorkbook(filePath, supplierName) {
  const extension = path.extname(filePath).toLowerCase();
  const workbook =
    extension === ".csv"
      ? XLSX.read(fs.readFileSync(filePath, "utf8").replace(/^\uFEFF/, ""), {
          type: "string",
          cellDates: true,
          codepage: 65001
        })
      : XLSX.readFile(filePath, { cellDates: true });
  if (!workbook.SheetNames.length) {
    throw new Error("文件中没有可读取的工作表");
  }

  const candidates = workbook.SheetNames.map((sheetName) =>
    parseWorksheet(workbook.Sheets[sheetName], supplierName, sheetName)
  );

  const bestCandidate = candidates.sort((a, b) => {
    if (b.score !== a.score) {
      return b.score - a.score;
    }
    return b.parsedCount - a.parsedCount;
  })[0];

  if (!bestCandidate || bestCandidate.parsedCount === 0) {
    const summary = candidates
      .slice(0, 3)
      .map((candidate) => {
        const headers = candidate.detectedHeaders.length
          ? candidate.detectedHeaders.join("、")
          : "未识别到清晰表头";
        return `工作表「${candidate.sheetName}」表头候选：${headers}`;
      })
      .join("；");

    throw new Error(`没有识别到可导入的商品数据。${summary}`);
  }

  return {
    rowCount: bestCandidate.rowCount,
    parsedCount: bestCandidate.parsedCount,
    items: bestCandidate.items
  };
}

function parseWorksheet(worksheet, supplierName, sheetName) {
  const grid = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: false
  });

  const headerRowIndex = findHeaderRowIndex(grid);
  const headerRow = headerRowIndex >= 0 ? grid[headerRowIndex] : [];
  const detectedHeaders = headerRow
    .map((cell) => String(cell || "").trim())
    .filter(Boolean)
    .slice(0, 12);

  if (headerRowIndex < 0) {
    return {
      sheetName,
      rowCount: 0,
      parsedCount: 0,
      items: [],
      detectedHeaders,
      score: 0
    };
  }

  const rows = XLSX.utils.sheet_to_json(worksheet, {
    range: headerRowIndex,
    defval: "",
    raw: false
  });

  const items = [];

  for (const row of rows) {
    const productName = pickValue(row, headerAliases.productName);
    const sku = pickValue(row, headerAliases.sku);
    const explicitSpec = pickValue(row, headerAliases.spec);
    const color = pickValue(row, headerAliases.color);
    const size = pickValue(row, headerAliases.size);
    const spec = explicitSpec || [color, size].filter(Boolean).join(" / ");
    const unit = pickValue(row, headerAliases.unit);
    const priceValue = pickValue(row, headerAliases.price);
    const moqValue = pickValue(row, headerAliases.moq);
    const stockValue = pickValue(row, headerAliases.stock);

    const hasIdentity = Boolean(productName || sku);
    const price = parseNumber(priceValue);

    if (!hasIdentity) {
      continue;
    }

    items.push({
      supplierName,
      productName,
      sku,
      spec,
      unit,
      price,
      moq: parseNumber(moqValue),
      stock: parseNumber(stockValue),
      raw: row
    });
  }

  return {
    sheetName,
    rowCount: rows.length,
    parsedCount: items.length,
    items,
    detectedHeaders,
    score: scoreHeaderRow(headerRow) + items.length * 10
  };
}

function findHeaderRowIndex(grid) {
  const searchLimit = Math.min(grid.length, 30);
  let bestIndex = -1;
  let bestScore = 0;

  for (let index = 0; index < searchLimit; index += 1) {
    const score = scoreHeaderRow(grid[index]);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  }

  return bestScore >= 16 ? bestIndex : -1;
}

function scoreHeaderRow(row) {
  if (!Array.isArray(row) || !row.length) {
    return 0;
  }

  const detectedTypes = new Set();
  let matchedCells = 0;
  let nonEmptyCells = 0;

  for (const cell of row) {
    if (String(cell || "").trim()) {
      nonEmptyCells += 1;
    }

    const types = matchHeaderTypes(cell);
    if (types.length) {
      matchedCells += 1;
    }

    for (const type of types) {
      detectedTypes.add(type);
    }
  }

  let score = matchedCells * 4 + detectedTypes.size * 5;
  if (detectedTypes.has("productName") || detectedTypes.has("sku")) {
    score += 12;
  }
  if (detectedTypes.has("price")) {
    score += 8;
  }
  if (nonEmptyCells >= 3) {
    score += 2;
  }

  return score;
}

function pickValue(row, aliases) {
  for (const [key, value] of Object.entries(row)) {
    const normalizedKey = normalizeHeader(key);
    if (aliases.some((alias) => headerMatches(normalizedKey, alias))) {
      return String(value || "").trim();
    }
  }

  return "";
}

function matchHeaderTypes(value) {
  const normalized = normalizeHeader(value);
  if (!normalized) {
    return [];
  }

  return Object.entries(headerAliases)
    .filter(([, aliases]) => aliases.some((alias) => headerMatches(normalized, alias)))
    .map(([type]) => type);
}

function headerMatches(normalizedValue, alias) {
  const normalizedAlias = normalizeHeader(alias);
  if (!normalizedValue || !normalizedAlias) {
    return false;
  }

  if (normalizedValue === normalizedAlias) {
    return true;
  }

  if (normalizedAlias.length >= 4 && normalizedValue.includes(normalizedAlias)) {
    return true;
  }

  if (normalizedValue.length >= 4 && normalizedAlias.includes(normalizedValue)) {
    return true;
  }

  return false;
}

function normalizeHeader(value) {
  return String(value || "")
    .replace(/^\uFEFF/, "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()（）【】\[\]_:：-]/g, "");
}

function normalizeText(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[^\p{L}\p{N}]+/gu, "");
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  const normalized = String(value).replace(/,/g, "").match(/-?\d+(\.\d+)?/);
  if (!normalized) {
    return null;
  }

  return Number(normalized[0]);
}

function buildComparisonGroups(products, groupBy) {
  const groups = new Map();

  for (const product of products) {
    const skuKey = normalizeText(product.sku);
    const nameKey = `${normalizeText(product.product_name)}::${normalizeText(product.spec)}`;

    let groupKey = nameKey;
    if (groupBy === "sku") {
      groupKey = skuKey || nameKey;
    } else if (groupBy === "smart") {
      groupKey = skuKey || nameKey;
    }

    if (!groupKey) {
      continue;
    }

    if (!groups.has(groupKey)) {
      groups.set(groupKey, {
        groupKey,
        productName: product.product_name || "未命名商品",
        sku: product.sku || "",
        spec: product.spec || "",
        unit: product.unit || "",
        offerCount: 0,
        minPrice: null,
        maxPrice: null,
        priceSpread: null,
        offers: []
      });
    }

    const group = groups.get(groupKey);
    group.productName = group.productName === "未命名商品" ? product.product_name || group.productName : group.productName;
    group.sku = group.sku || product.sku || "";
    group.spec = group.spec || product.spec || "";
    group.unit = group.unit || product.unit || "";
    group.offers.push({
      supplierName: product.supplier_name,
      price: product.price,
      moq: product.moq,
      stock: product.stock,
      uploadedAt: product.uploaded_at
    });
  }

  return Array.from(groups.values())
    .map((group) => {
      group.offers.sort((a, b) => {
        if (a.price === null && b.price === null) {
          return 0;
        }
        if (a.price === null) {
          return 1;
        }
        if (b.price === null) {
          return -1;
        }
        return a.price - b.price;
      });

      const prices = group.offers.map((offer) => offer.price).filter((price) => price !== null);
      group.offerCount = group.offers.length;
      group.minPrice = prices.length ? Math.min(...prices) : null;
      group.maxPrice = prices.length ? Math.max(...prices) : null;
      group.priceSpread =
        group.minPrice !== null && group.maxPrice !== null
          ? Number((group.maxPrice - group.minPrice).toFixed(2))
          : null;

      return group;
    })
    .sort((a, b) => {
      if (b.offerCount !== a.offerCount) {
        return b.offerCount - a.offerCount;
      }
      return (b.priceSpread || 0) - (a.priceSpread || 0);
    });
}

async function organizeProductsWithAi({ collection, keyword, sourceRows }) {
  const totalRows = sourceRows.length;
  const maxRows = 20;
  const limitedRows = sourceRows.slice(0, maxRows);
  const truncated = totalRows > maxRows;

  if (!limitedRows.length) {
    return {
      collection: formatCollection(collection),
      keyword,
      total: 0,
      truncated: false,
      model: aiRuntimeConfig.model,
      summary: "当前没有可整理的货盘数据。",
      columns: [],
      items: []
    };
  }

  const promptPayload = limitedRows.map((row) => ({
    id: row.id,
    supplier_name: row.supplierName,
    raw: row.raw
  }));

  const aiResult = await callDashScopeJson({
    systemPrompt: [
      "你是一个供应商货盘整理助手。",
      "你的任务是根据原始 Excel 行数据，自动识别这一批货盘最适合展示的字段，并整理成统一结构。",
      "必须输出 JSON，不要输出任何额外说明。",
      "不要编造不存在的值；缺失时返回 null。",
      "请优先提取有助于选品和比价的字段，例如商品名称、款号、颜色、尺码、价格、品牌、品类、库存、起订量。",
      "columns 最多 8 个；key 使用 snake_case 英文；label 使用中文。"
    ].join(" "),
    userPrompt: JSON.stringify(
      {
        task: "请根据这些货盘行数据输出 JSON，结构为 { summary, columns, rows, warnings }。rows 中每一项必须保留 id，并在 values 中按 columns.key 输出整理后的字段值。",
        collection_name: collection.name,
        collection_description: collection.description || "",
        keyword: keyword || null,
        row_count: limitedRows.length,
        rows: promptPayload
      },
      null,
      2
    )
  });

  const columns = normalizeAiColumns(aiResult.columns);
  const rowsById = new Map(
    (Array.isArray(aiResult.rows) ? aiResult.rows : []).map((row) => [String(row.id), row])
  );

  const items = limitedRows.map((sourceRow) => {
    const aiRow = rowsById.get(String(sourceRow.id)) || {};
    const values = {};

    for (const column of columns) {
      values[column.key] = stringifyAiCell(aiRow.values?.[column.key]);
    }

    return {
      id: sourceRow.id,
      supplierName: sourceRow.supplierName,
      uploadedAt: sourceRow.uploadedAt,
      values,
      raw: sourceRow.raw
    };
  });

  return {
    collection: formatCollection(collection),
    keyword,
    total: totalRows,
    visibleCount: items.length,
    truncated,
    model: aiRuntimeConfig.model,
    summary:
      stringifyAiCell(aiResult.summary) ||
      `已用 ${aiRuntimeConfig.model} 自动整理 ${items.length} 条货盘记录。`,
    warnings: Array.isArray(aiResult.warnings) ? aiResult.warnings.map(stringifyAiCell).filter(Boolean) : [],
    columns,
    items
  };
}

async function summarizeCatalogForStorage({ collection, supplierName, originalFilename, parsed }) {
  const fallback = buildCatalogFallbackSummary({ parsed });

  if (!aiRuntimeConfig.apiKey) {
    return fallback;
  }

  try {
    const sourceRows = parsed.items.slice(0, 24).map((item, index) => ({
      index: index + 1,
      productName: item.productName,
      sku: item.sku,
      spec: item.spec,
      price: item.price,
      unit: item.unit,
      raw: compactRawRow(item.raw)
    }));

    const aiResult = await callDashScopeJson({
      systemPrompt: [
        "你是一个货盘归档助手。",
        "请根据供应商上传的 Excel/CSV 货盘内容，提取适合在列表中展示的摘要信息。",
        "必须输出 JSON，不要输出任何额外说明。",
        "返回结构必须是 { category_summary, product_names_summary, catalog_overview }。",
        "category_summary 和 product_names_summary 都返回字符串数组，最多 6 项。",
        "catalog_overview 返回一段不超过 60 字的中文摘要。"
      ].join(" "),
      userPrompt: JSON.stringify(
        {
          collectionName: collection.name,
          supplierName,
          originalFilename,
          parsedCount: parsed.parsedCount,
          rowCount: parsed.rowCount,
          rows: sourceRows
        },
        null,
        2
      )
    });

    const categorySummary = normalizeSummaryItems(aiResult.category_summary, 6);
    const productNamesSummary = normalizeSummaryItems(aiResult.product_names_summary, 6);
    const catalogOverview = String(aiResult.catalog_overview || "").trim();

    return {
      categorySummary: categorySummary.length ? categorySummary.join(" / ") : fallback.categorySummary,
      productNamesSummary: productNamesSummary.length
        ? productNamesSummary.join(" / ")
        : fallback.productNamesSummary,
      catalogOverview: catalogOverview || fallback.catalogOverview,
      analysisSource: "ai"
    };
  } catch (_error) {
    return fallback;
  }
}

function buildCatalogFallbackSummary({ parsed }) {
  const categories = inferCategoryLabels(
    parsed.items.map((item) => ({
      raw_json: JSON.stringify(item.raw),
      product_name: item.productName
    }))
  );
  const productNames = inferProductNames(
    parsed.items.map((item) => ({
      product_name: item.productName
    }))
  );

  return {
    categorySummary: categories.length ? categories.join(" / ") : "未识别",
    productNamesSummary: productNames.length ? productNames.join(" / ") : "未识别",
    catalogOverview: `共导入 ${parsed.parsedCount} 个商品，已生成基础摘要`,
    analysisSource: "rule"
  };
}

async function summarizeCollectionWithAi({ collection, rows }) {
  const totalProducts = rows.length;
  const supplierMap = new Map();
  let pricedCount = 0;
  let stockCount = 0;
  const userProfile = getCurrentQjlProfileContext();

  for (const row of rows) {
    if (row.price !== null && row.price !== undefined) {
      pricedCount += 1;
    }
    if (row.stock !== null && row.stock !== undefined) {
      stockCount += 1;
    }

    if (!supplierMap.has(row.supplier_name)) {
      supplierMap.set(row.supplier_name, {
        supplierName: row.supplier_name,
        productCount: 0,
        pricedCount: 0,
        sampleProducts: []
      });
    }

    const supplier = supplierMap.get(row.supplier_name);
    supplier.productCount += 1;
    if (row.price !== null && row.price !== undefined) {
      supplier.pricedCount += 1;
    }
    if (supplier.sampleProducts.length < 5) {
      supplier.sampleProducts.push({
        productName: row.product_name,
        sku: row.sku,
        spec: row.spec,
        price: row.price
      });
    }
  }

  const supplierSummary = Array.from(supplierMap.values());
  const sampleRows = rows.slice(0, 120).map((row) => ({
    supplierName: row.supplier_name,
    productName: row.product_name,
    sku: row.sku,
    spec: row.spec,
    unit: row.unit,
    price: row.price,
    moq: row.moq,
    stock: row.stock
  }));

  const aiResult = await callDashScopeJson({
    systemPrompt: [
      "你是一个货盘分析助手，负责根据多家供应商上传的货盘数据做汇总分析。",
      "请输出 JSON，不要输出额外解释。",
      "目标是帮助选品人快速了解这批货盘的规模、质量、可比价性和需要补充的信息。",
      "如果给出了登录用户画像，请结合该用户的经营阶段、偏好品类和选品关注点来判断这批货盘的匹配度。",
      "返回结构必须是 { overview, highlights, supplier_insights, next_actions }。",
      "overview 是一段简短中文总结；highlights、supplier_insights、next_actions 都是字符串数组。"
    ].join(" "),
    userPrompt: JSON.stringify(
      {
        collection: {
          name: collection.name,
          description: collection.description || ""
        },
        stats: {
          total_products: totalProducts,
          supplier_count: supplierSummary.length,
          priced_count: pricedCount,
          stock_count: stockCount
        },
        user_profile: userProfile,
        suppliers: supplierSummary,
        sample_rows: sampleRows
      },
      null,
      2
    )
  });

  return {
    collection: formatCollection(collection),
    model: aiRuntimeConfig.model,
    stats: {
      totalProducts,
      supplierCount: supplierSummary.length,
      pricedCount,
      stockCount
    },
    overview:
      stringifyAiCell(aiResult.overview) ||
      `本批次共导入 ${totalProducts} 条货盘，来自 ${supplierSummary.length} 家供应商。`,
    highlights: Array.isArray(aiResult.highlights) ? aiResult.highlights.map(stringifyAiCell).filter(Boolean) : [],
    supplierInsights: Array.isArray(aiResult.supplier_insights)
      ? aiResult.supplier_insights.map(stringifyAiCell).filter(Boolean)
      : [],
    nextActions: Array.isArray(aiResult.next_actions)
      ? aiResult.next_actions.map(stringifyAiCell).filter(Boolean)
      : []
  };
}

function normalizeSummaryItems(items, limit = 6) {
  if (!Array.isArray(items)) {
    return [];
  }

  return items.map(stringifyAiCell).filter(Boolean).slice(0, limit);
}

function buildUploadChatContext({ collection, uploads, products }) {
  const byUpload = new Map();

  for (const product of products) {
    if (!byUpload.has(product.upload_id)) {
      byUpload.set(product.upload_id, []);
    }
    byUpload.get(product.upload_id).push(product);
  }

  const uploadSummaries = uploads.map((upload) => {
    const rows = byUpload.get(upload.id) || [];
    const categoryLabels = inferCategoryLabels(rows);
    const prices = rows.map((row) => row.price).filter((value) => value !== null && value !== undefined);
    const sampleRows = rows.slice(0, 8).map((row) => ({
      supplierName: row.supplier_name,
      productName: row.product_name,
      sku: row.sku,
      spec: row.spec,
      unit: row.unit,
      price: row.price,
      moq: row.moq,
      stock: row.stock,
      raw: compactRawRow(safeParseJson(row.raw_json))
    }));

    return {
      uploadId: upload.id,
      supplierName: upload.supplier_name,
      catalogName: upload.original_filename,
      uploadedAt: upload.uploaded_at,
      productCount: Number(upload.parsed_count || rows.length || 0),
      categoryLabels,
      priceRange: prices.length
        ? {
            min: Math.min(...prices),
            max: Math.max(...prices)
          }
        : null,
      sampleRows
    };
  });

  const totalProducts = uploadSummaries.reduce((sum, item) => sum + item.productCount, 0);
  const allCategories = uniqueStrings(uploadSummaries.flatMap((item) => item.categoryLabels)).slice(0, 12);
  const supplierCount = uniqueStrings(uploadSummaries.map((item) => item.supplierName)).length;

  return {
    collection: {
      name: collection.name,
      description: collection.description || ""
    },
    stats: {
      selected_uploads: uploadSummaries.length,
      total_products: totalProducts,
      supplier_count: supplierCount,
      categories: allCategories
    },
    uploads: uploadSummaries
  };
}

function buildUploadChatContextAcrossCollections({ uploads, products }) {
  const byUpload = new Map();

  for (const product of products) {
    if (!byUpload.has(product.upload_id)) {
      byUpload.set(product.upload_id, []);
    }
    byUpload.get(product.upload_id).push(product);
  }

  const uploadSummaries = uploads.map((upload) => {
    const rows = byUpload.get(upload.id) || [];
    const categoryLabels = inferCategoryLabels(rows);
    const prices = rows.map((row) => row.price).filter((value) => value !== null && value !== undefined);
    const sampleRows = rows.slice(0, 8).map((row) => ({
      supplierName: row.supplier_name,
      productName: row.product_name,
      sku: row.sku,
      spec: row.spec,
      unit: row.unit,
      price: row.price,
      moq: row.moq,
      stock: row.stock,
      raw: compactRawRow(safeParseJson(row.raw_json))
    }));

    return {
      uploadId: upload.id,
      batchName: upload.collection_name,
      batchDescription: upload.collection_description || "",
      supplierName: upload.supplier_name,
      catalogName: upload.original_filename,
      uploadedAt: upload.uploaded_at,
      productCount: Number(upload.parsed_count || rows.length || 0),
      categoryLabels,
      priceRange: prices.length
        ? {
            min: Math.min(...prices),
            max: Math.max(...prices)
          }
        : null,
      sampleRows
    };
  });

  const totalProducts = uploadSummaries.reduce((sum, item) => sum + item.productCount, 0);
  const allCategories = uniqueStrings(uploadSummaries.flatMap((item) => item.categoryLabels)).slice(0, 12);
  const supplierCount = uniqueStrings(uploadSummaries.map((item) => item.supplierName)).length;
  const batchNames = uniqueStrings(uploadSummaries.map((item) => item.batchName));

  return {
    collection: {
      name: "历史货盘批次",
      description: `已选 ${batchNames.length} 个批次`
    },
    stats: {
      selected_uploads: uploadSummaries.length,
      selected_batches: batchNames.length,
      total_products: totalProducts,
      supplier_count: supplierCount,
      categories: allCategories
    },
    uploads: uploadSummaries
  };
}

async function chatAboutUploadsWithAi({ collection, message, history, context }) {
  const normalizedHistory = history
    .filter((item) => item && (item.role === "user" || item.role === "assistant"))
    .map((item) => ({
      role: item.role,
      content: String(item.content || "").trim()
    }))
    .filter((item) => item.content)
    .slice(-8);

  const systemPrompt = [
    "你是一个供应商货盘分析助手。",
    "用户会勾选若干份货盘，然后和你继续追问。",
    "请基于给定的货盘上下文回答，重点关注价格结构、品类覆盖、供应商差异、风险点、缺失信息和下一步建议。",
    "如果上下文里包含登录用户的群接龙画像，请结合该用户的品类偏好、经营阶段和选品关注点来给出更贴近业务的建议。",
    "不要假装看过上下文之外的数据；如果信息不足，要明确说出缺口。",
    "回答使用中文，尽量结构化、简洁、可执行。",
    "适合时请使用 Markdown 小标题、项目符号和表格来组织结果，尤其是涉及多供应商价格、优缺点对比、风险清单时。"
  ].join(" ");

  const userProfile = getCurrentQjlProfileContext();

  const contextMessage = JSON.stringify(
    {
      task: "以下是当前勾选货盘的上下文，请据此回答用户问题。",
      collection: {
        name: collection.name,
        description: collection.description || ""
      },
      userProfile,
      context
    },
    null,
    2
  );

  return callDashScopeText({
    systemPrompt,
    contextMessage,
    history: normalizedHistory,
    userMessage: message
  });
}

function normalizeAiColumns(columns) {
  const fallback = [
    { key: "product_name", label: "商品名称" },
    { key: "sku_or_style_code", label: "款号 / SKU" },
    { key: "spec_or_variant", label: "规格 / 变体" }
  ];

  if (!Array.isArray(columns) || !columns.length) {
    return fallback;
  }

  const normalized = [];
  const used = new Set();

  for (const column of columns) {
    const key = String(column?.key || "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9_]+/g, "_")
      .replace(/^_+|_+$/g, "");
    const label = String(column?.label || "").trim();

    if (!key || !label || used.has(key)) {
      continue;
    }

    used.add(key);
    normalized.push({ key, label });
  }

  return normalized.length ? normalized.slice(0, 8) : fallback;
}

function stringifyAiCell(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  if (typeof value === "string") {
    return value.trim();
  }

  if (Array.isArray(value)) {
    return value.map((item) => stringifyAiCell(item)).filter(Boolean).join(" / ");
  }

  if (typeof value === "object") {
    return JSON.stringify(value);
  }

  return String(value);
}

async function callDashScopeJson({ systemPrompt, userPrompt }) {
  const baseUrl = aiRuntimeConfig.baseUrl.replace(/\/$/, "");
  const response = await fetch(`${baseUrl}/chat/completions`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${aiRuntimeConfig.apiKey}`
    },
    body: JSON.stringify({
      model: aiRuntimeConfig.model,
      messages: [
        {
          role: "system",
          content: `${systemPrompt} 请只返回 JSON。`
        },
        {
          role: "user",
          content: `请以 JSON 输出整理结果。\n${userPrompt}`
        }
      ],
      response_format: {
        type: "json_object"
      },
      enable_thinking: false,
      temperature: 0.2
    })
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`百炼调用失败：${response.status} ${errorText}`);
  }

  const data = await response.json();
  const content = data?.choices?.[0]?.message?.content;

  if (!content) {
    throw new Error("百炼没有返回可解析的内容");
  }

  try {
    return JSON.parse(content);
  } catch (_error) {
    throw new Error("百炼返回的内容不是有效 JSON");
  }
}

async function callDashScopeText({ systemPrompt, contextMessage, history, userMessage }) {
  const baseUrl = aiRuntimeConfig.baseUrl.replace(/\/$/, "");
  const messages = [
    {
      role: "system",
      content: systemPrompt
    },
    {
      role: "user",
      content: contextMessage
    },
    {
      role: "assistant",
      content: "我已收到当前勾选货盘的上下文，可以继续分析。"
    },
    ...history,
    {
      role: "user",
      content: userMessage
    }
  ];

  const response = await fetch(`${baseUrl}/chat/completions`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${aiRuntimeConfig.apiKey}`
    },
    body: JSON.stringify({
      model: aiRuntimeConfig.model,
      messages,
      enable_thinking: false,
      temperature: 0.3
    })
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`百炼调用失败：${response.status} ${errorText}`);
  }

  const data = await response.json();
  const content = data?.choices?.[0]?.message?.content;

  if (!content) {
    throw new Error("百炼没有返回可解析的内容");
  }

  return String(content).trim();
}

function uniqueStrings(values) {
  const used = new Set();
  const result = [];

  for (const value of values || []) {
    const text = String(value || "").trim();
    if (!text) {
      continue;
    }

    const normalized = normalizeText(text);
    if (!normalized || used.has(normalized)) {
      continue;
    }

    used.add(normalized);
    result.push(text);
  }

  return result;
}

function renderUploadPage(collection) {
  const escapedName = escapeHtml(collection.name);
  const escapedDescription = escapeHtml(collection.description || "");

  return `
    <!DOCTYPE html>
    <html lang="zh-CN">
      <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>${escapedName} - 供应商货盘上传</title>
        <style>
          :root {
            color-scheme: light;
            --bg: #f5efe3;
            --card: rgba(255, 252, 246, 0.92);
            --ink: #1e2b24;
            --muted: #5b665f;
            --line: rgba(30, 43, 36, 0.12);
            --brand: #14532d;
            --brand-soft: #dbe9dc;
            --accent: #ca8a04;
          }
          * { box-sizing: border-box; }
          body {
            margin: 0;
            min-height: 100vh;
            font-family: "PingFang SC", "Noto Sans SC", "Microsoft YaHei", sans-serif;
            color: var(--ink);
            background:
              radial-gradient(circle at top right, rgba(202, 138, 4, 0.18), transparent 30%),
              radial-gradient(circle at bottom left, rgba(20, 83, 45, 0.18), transparent 32%),
              var(--bg);
            display: grid;
            place-items: center;
            padding: 24px;
          }
          .panel {
            width: min(680px, 100%);
            background: var(--card);
            border: 1px solid var(--line);
            border-radius: 24px;
            box-shadow: 0 18px 70px rgba(30, 43, 36, 0.12);
            padding: 28px;
            backdrop-filter: blur(8px);
          }
          .eyebrow {
            display: inline-flex;
            padding: 6px 12px;
            border-radius: 999px;
            background: var(--brand-soft);
            color: var(--brand);
            font-size: 13px;
            font-weight: 700;
          }
          h1 {
            margin: 16px 0 10px;
            font-size: clamp(28px, 6vw, 42px);
            line-height: 1.1;
          }
          p {
            color: var(--muted);
            line-height: 1.7;
          }
          form {
            margin-top: 24px;
            display: grid;
            gap: 16px;
          }
          label {
            display: grid;
            gap: 8px;
            font-weight: 600;
          }
          input, button {
            font: inherit;
          }
          input[type="text"],
          input[type="file"] {
            width: 100%;
            padding: 14px 16px;
            border-radius: 14px;
            border: 1px solid var(--line);
            background: white;
          }
          button {
            border: none;
            border-radius: 14px;
            padding: 14px 18px;
            background: linear-gradient(135deg, var(--brand), #1f7a43);
            color: white;
            font-weight: 700;
            cursor: pointer;
          }
          button:disabled {
            opacity: 0.7;
            cursor: wait;
          }
          .helper {
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            margin-top: 12px;
            font-size: 14px;
          }
          .helper a {
            color: var(--brand);
          }
          .result {
            margin-top: 20px;
            padding: 14px 16px;
            border-radius: 14px;
            display: none;
            line-height: 1.6;
          }
          .result.ok {
            display: block;
            background: #e7f7eb;
            color: #12522d;
          }
          .result.error {
            display: block;
            background: #fff1f2;
            color: #9f1239;
          }
        </style>
      </head>
      <body>
        <section class="panel">
          <div class="eyebrow">供应商货盘上传入口</div>
          <h1>${escapedName}</h1>
          <p>${escapedDescription || "请填写供应商名称，并上传最新的 Excel / CSV 货盘。系统会自动识别商品和价格，方便后续统一比价。"}</p>

          <form id="upload-form">
            <label>
              供应商名称
              <input type="text" name="supplierName" placeholder="例如：山东果园直发" required />
            </label>

            <label>
              上传货盘文件
              <input type="file" name="file" accept=".xlsx,.xls,.csv" required />
            </label>

            <button type="submit">提交货盘</button>
          </form>

          <div class="helper">
            <span>支持格式：.xlsx、.xls、.csv，单个文件不超过 50MB</span>
            <a href="/template.xlsx">下载模板</a>
          </div>

          <div id="result" class="result"></div>
        </section>

        <script>
          const form = document.getElementById("upload-form");
          const result = document.getElementById("result");
          const button = form.querySelector("button");

          form.addEventListener("submit", async (event) => {
            event.preventDefault();
            result.className = "result";
            result.textContent = "";
            button.disabled = true;
            button.textContent = "上传中...";

            try {
              const formData = new FormData(form);
              const response = await fetch("/api/public/upload/${collection.upload_token}", {
                method: "POST",
                body: formData
              });
              const data = await response.json();

              if (!response.ok) {
                throw new Error(data.error || "上传失败");
              }

              result.className = "result ok";
              result.textContent = "上传成功：已接收 " + data.rowCount + " 行，识别并导入 " + data.parsedCount + " 条商品记录。";
              form.reset();
            } catch (error) {
              result.className = "result error";
              result.textContent = error.message || "上传失败，请稍后重试。";
            } finally {
              button.disabled = false;
              button.textContent = "提交货盘";
            }
          });
        </script>
      </body>
    </html>
  `;
}

function renderMissingLinkPage() {
  return `
    <!DOCTYPE html>
    <html lang="zh-CN">
      <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>链接不可用</title>
        <style>
          body {
            margin: 0;
            min-height: 100vh;
            display: grid;
            place-items: center;
            background: #f7f1e5;
            color: #322d25;
            font-family: "PingFang SC", "Noto Sans SC", "Microsoft YaHei", sans-serif;
          }
          .card {
            width: min(520px, calc(100% - 40px));
            background: white;
            border-radius: 24px;
            padding: 28px;
            box-shadow: 0 16px 60px rgba(50, 45, 37, 0.12);
          }
          h1 { margin-top: 0; }
          p { line-height: 1.7; color: #6b5f52; }
        </style>
      </head>
      <body>
        <div class="card">
          <h1>这个上传链接不可用了</h1>
          <p>可能是链接填写错误，或者该批次已经被删除。请联系收集货盘的人重新发送正确链接。</p>
        </div>
      </body>
    </html>
  `;
}

function normalizeExternalSearchPayload(payload, platformLabel) {
  const data = payload?.result || payload?.data || payload?.items || payload?.item || payload || {};
  const items = firstArray(
    data.items?.item,
    data.items,
    data.items,
    data.item,
    data.list,
    data.auctions,
    data.products,
    payload?.items?.item,
    payload?.items
  );
  const total =
    Number(
      data.total_results ||
        data.real_total_results ||
        data.total ||
        data.search_count ||
        payload?.items?.total_results ||
        payload?.items?.real_total_results ||
        payload?.total ||
        items.length
    ) || items.length;

  return {
    total,
    items: items.map((item) => normalizeExternalSearchItem(item, platformLabel)).filter(Boolean)
  };
}

function normalizeExternalSearchItem(item, platformLabel) {
  if (!item || typeof item !== "object") {
    return null;
  }

  const title = firstString(item.title, item.name, item.item_title, item.raw_title, item.subject);
  const price = firstString(
    item.price,
    item.promotion_price,
    item.zk_final_price,
    item.orginal_price,
    item.original_price,
    item.sale_price
  );
  const detailUrl = firstString(item.detail_url, item.url, item.item_url, item.click_url);
  const shopName = firstString(item.shop_name, item.seller_name, item.nick, item.store_name);
  const itemId = firstString(item.num_iid, item.itemid, item.item_id, item.sku_id, item.id);
  const location = firstString(item.provcity, item.location, item.area, item.city);
  const sales = firstString(item.sales, item.sale_num, item.sold, item.volume);
  const imageUrl = normalizeExternalImageUrl(
    firstString(
      item.pic_url,
      item.pict_url,
      item.picUrl,
      item.image_url,
      item.imageUrl,
      item.main_pic,
      item.mainPic,
      item.img,
      item.image,
      item.thumbnail_url,
      item.thumbnailUrl,
      item.thumbnail,
      item.small_images?.string?.[0],
      item.small_images?.[0],
      item.pics?.[0]
    )
  );

  return {
    title: title || "未命名商品",
    priceText: price ? `¥${price}` : "-",
    salesText: sales ? `销量 ${sales}` : "外部价格",
    detailUrl,
    imageUrl,
    shopName: shopName || "未知店铺",
    itemId: itemId || "",
    location: location || "",
    platformLabel
  };
}

function normalizeExternalImageUrl(value) {
  const imageUrl = String(value || "").trim();
  if (!imageUrl) {
    return "";
  }

  if (imageUrl.startsWith("//")) {
    return `https:${imageUrl}`;
  }

  return imageUrl;
}

function filterExternalItemsByMatchMode({ items, keyword, matchMode }) {
  if (matchMode !== "brand_style") {
    return items;
  }

  const signature = buildExternalKeywordSignature(keyword);
  if (!signature.brandTokens.length && !signature.styleTokens.length) {
    return items;
  }

  return items.filter((item) => matchesBrandAndStyle(item?.title || "", signature));
}

function buildExternalKeywordSignature(keyword) {
  const normalized = String(keyword || "").trim().toUpperCase();
  const styleTokens = uniqueStrings(normalized.match(/[A-Z]+[A-Z0-9-]*\d+[A-Z0-9-]*/g) || []);

  let brandSource = normalized;
  for (const token of styleTokens) {
    brandSource = brandSource.replaceAll(token, " ");
  }

  const stopWords = [
    "手持小风扇",
    "小风扇",
    "风扇",
    "口红款",
    "口红",
    "手持",
    "便携",
    "迷你",
    "折叠",
    "款",
    "新款",
    "新品",
    "充电",
    "高速",
    "静音",
    "制冷",
    "喷雾",
    "数显",
    "桌面",
    "涡轮",
    "无级变速",
    "USB"
  ];
  for (const word of stopWords) {
    brandSource = brandSource.replaceAll(word, " ");
  }

  const brandTokens = uniqueStrings((brandSource.match(/[\u4E00-\u9FFF]{2,8}|[A-Z]{2,}/g) || []).filter((token) => token !== "USB"));

  return {
    brandTokens,
    styleTokens
  };
}

function matchesBrandAndStyle(title, signature) {
  const normalizedTitle = String(title || "").trim().toUpperCase();
  const brandOk =
    !signature.brandTokens.length ||
    signature.brandTokens.some((token) => normalizedTitle.includes(token));
  const styleOk =
    !signature.styleTokens.length ||
    signature.styleTokens.some((token) => normalizedTitle.includes(token));

  return brandOk && styleOk;
}

function firstArray(...candidates) {
  for (const candidate of candidates) {
    if (Array.isArray(candidate)) {
      return candidate;
    }
  }
  return [];
}

function firstString(...candidates) {
  for (const candidate of candidates) {
    const value = String(candidate ?? "").trim();
    if (value) {
      return value;
    }
  }
  return "";
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
