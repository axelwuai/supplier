const state = {
  collections: [],
  activeCollection: null,
  uploadBatches: [],
  selectedUploadIds: [],
  chatMessages: []
};

const workspaceStage = document.querySelector(".workspace-stage");
const createForm = document.getElementById("create-form");
const workspaceLayout = document.querySelector(".workspace-layout");
const workspaceResizer = document.getElementById("workspace-resizer");
const settingsDetails = document.getElementById("settings-details");
const settingsSummaryText = document.getElementById("settings-summary-text");
const llmConfigForm = document.getElementById("llm-config-form");
const clearLlmConfigButton = document.getElementById("clear-llm-config");
const llmConfigStatus = document.getElementById("llm-config-status");
const qjlAuthSummary = document.getElementById("qjl-auth-summary");
const qjlAuthStatus = document.getElementById("qjl-auth-status");
const qjlLoginLink = document.getElementById("qjl-login-link");
const qjlProfilePreview = document.getElementById("qjl-profile-preview");
const qjlProfileName = document.getElementById("qjl-profile-name");
const qjlProfileMeta = document.getElementById("qjl-profile-meta");
const qjlHomepageMenu = document.getElementById("qjl-homepage-menu");
const qjlProfileSummary = document.getElementById("qjl-profile-summary");
const qjlProfileTags = document.getElementById("qjl-profile-tags");
const qjlLoginStartButton = document.getElementById("qjl-login-start");
const qjlLoginCheckButton = document.getElementById("qjl-login-check");
const qjlRefreshProfileButton = document.getElementById("qjl-refresh-profile");
const qjlLogoutButton = document.getElementById("qjl-logout");
const toggleUploadEntryButton = document.getElementById("toggle-upload-entry");
const closeUploadEntryButton = document.getElementById("close-upload-entry");
const uploadEntryPanel = document.getElementById("upload-entry-panel");
const createResult = document.getElementById("create-result");
const collectionList = document.getElementById("collection-list");
const emptyState = document.getElementById("empty-state");
const detailContent = document.getElementById("detail-content");
const detailName = document.getElementById("detail-name");
const detailDescription = document.getElementById("detail-description");
const detailBadge = document.getElementById("detail-badge");
const detailUploadCount = document.getElementById("detail-upload-count");
const detailProductCount = document.getElementById("detail-product-count");
const detailAiStatus = document.getElementById("detail-ai-status");
const uploadTable = document.getElementById("upload-table");
const refreshUploads = document.getElementById("refresh-uploads");
const toggleAllUploads = document.getElementById("toggle-all-uploads");
const selectedUploadCount = document.getElementById("selected-upload-count");
const sendSelectedToAiButton = document.getElementById("send-selected-to-ai");
const compareBody = document.getElementById("compare-body");
const compareForm = document.getElementById("compare-form");
const compareResult = document.getElementById("compare-result");
const toggleCompareFullscreenButton = document.getElementById("toggle-compare-fullscreen");
const compareFullscreenModal = document.getElementById("compare-fullscreen-modal");
const compareFullscreenCard = document.getElementById("compare-fullscreen-card");
const compareFullscreenTarget = document.getElementById("compare-fullscreen-target");
const closeCompareFullscreenButton = document.getElementById("close-compare-fullscreen");
const clearChatButton = document.getElementById("clear-chat");
const chatSelectionStatus = document.getElementById("chat-selection-status");
const chatMessages = document.getElementById("chat-messages");
const chatForm = document.getElementById("chat-form");
const chatInput = document.getElementById("chat-input");
const productForm = document.getElementById("product-form");
const refreshProducts = document.getElementById("refresh-products");
const productSummary = document.getElementById("product-summary");
const productTable = document.getElementById("product-table");
const aiStatus = document.getElementById("ai-status");

const aiState = {
  enabled: false,
  provider: "dashscope-compatible",
  model: "qwen-plus-latest",
  baseUrl: "",
  source: "missing",
  hasApiKey: false,
  apiKeyHint: "",
  loaded: false
};

const qjlState = {
  loggedIn: false,
  pendingLogin: false,
  loginUrl: "",
  account: null,
  loaded: false,
  homepageSwitcherOpen: false
};

const resizeState = {
  dragging: false
};

const compareFullscreenState = {
  open: false,
  placeholder: null
};

init();

async function init() {
  initWorkspaceResizer();
  initUploadEntry();
  initCompareFullscreen();
  initPanelCollapse();
  await loadAiStatus();
  await loadQjlAccount();
  await loadCollections();
  await renderUploads();

  const collectionId = new URLSearchParams(window.location.search).get("collectionId");
  if (collectionId) {
    const found = state.collections.find((item) => item.id === collectionId);
    if (found) {
      await selectCollection(found);
      return;
    }
  }

  if (state.collections[0]) {
    await selectCollection(state.collections[0]);
    return;
  }

  renderChatPanel();
}

function initPanelCollapse() {
  const collapseButtons = document.querySelectorAll('.panel-collapse-button');

  collapseButtons.forEach(button => {
    button.addEventListener('click', function() {
      const panel = this.closest('.panel');
      const panelContent = panel.querySelector('.panel-content, .chat-workspace, #empty-state');

      if (panelContent) {
        const isCollapsed = panel.classList.contains('collapsed');

        if (isCollapsed) {
          // 展开面板
          panel.classList.remove('collapsed');
          panelContent.style.display = 'flex';
          this.innerHTML = '<svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor"><path d="M6 9l6 6 6-6z"/></svg>';
        } else {
          // 折叠面板
          panel.classList.add('collapsed');
          panelContent.style.display = 'none';
          this.innerHTML = '<svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor"><path d="M12 5v14"/></svg>';
        }
      }
    });
  });
}

function syncBodyModalState() {
  document.body.classList.toggle("modal-open", Boolean(document.querySelector(".modal-shell.is-open")));
}

function initUploadEntry() {
  if (!toggleUploadEntryButton || !uploadEntryPanel) {
    return;
  }

  const setUploadEntryOpen = (open) => {
    uploadEntryPanel.hidden = !open;
    uploadEntryPanel.setAttribute("aria-hidden", open ? "false" : "true");
    uploadEntryPanel.classList.toggle("is-open", open);
    toggleUploadEntryButton.setAttribute("aria-expanded", open ? "true" : "false");
    syncBodyModalState();
  };

  toggleUploadEntryButton.addEventListener("click", () => {
    const nextOpen = uploadEntryPanel.hidden;
    setUploadEntryOpen(nextOpen);

    if (nextOpen) {
      const firstInput = createForm?.querySelector("input, textarea");
      firstInput?.focus();
    }
  });

  closeUploadEntryButton?.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    setUploadEntryOpen(false);
  });

  uploadEntryPanel.addEventListener("click", (event) => {
    const target = event.target;
    if (target instanceof HTMLElement && target.dataset.closeUploadEntry === "true") {
      setUploadEntryOpen(false);
    }
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !uploadEntryPanel.hidden) {
      setUploadEntryOpen(false);
    }

    if (event.key === "Escape" && qjlState.homepageSwitcherOpen) {
      qjlState.homepageSwitcherOpen = false;
      renderQjlAuthState();
    }
  });

  document.addEventListener("click", (event) => {
    if (!qjlState.homepageSwitcherOpen || !qjlHomepageMenu || !qjlLoginStartButton) {
      return;
    }

    const target = event.target;
    if (!(target instanceof Node)) {
      return;
    }

    if (qjlHomepageMenu.contains(target) || qjlLoginStartButton.contains(target)) {
      return;
    }

    qjlState.homepageSwitcherOpen = false;
    renderQjlAuthState();
  });

  setUploadEntryOpen(false);
}

function initCompareFullscreen() {
  if (
    !toggleCompareFullscreenButton ||
    !compareBody ||
    !compareFullscreenModal ||
    !compareFullscreenCard ||
    !compareFullscreenTarget
  ) {
    return;
  }

  const setCompareFullscreenOpen = (open) => {
    compareFullscreenState.open = open;
    compareFullscreenModal.hidden = !open;
    compareFullscreenModal.setAttribute("aria-hidden", open ? "false" : "true");
    compareFullscreenModal.classList.toggle("is-open", open);
    toggleCompareFullscreenButton.setAttribute("aria-expanded", open ? "true" : "false");

    if (open) {
      if (!compareFullscreenState.placeholder) {
        compareFullscreenState.placeholder = document.createComment("compare-body-placeholder");
      }

      if (compareBody.parentNode !== compareFullscreenTarget) {
        compareBody.replaceWith(compareFullscreenState.placeholder);
        compareFullscreenTarget.appendChild(compareBody);
      }

      syncCompareFullscreenWidth();
      closeCompareFullscreenButton?.focus();
    } else {
      if (compareFullscreenState.placeholder?.parentNode) {
        compareFullscreenState.placeholder.replaceWith(compareBody);
      }
      compareFullscreenCard.style.removeProperty("width");
    }

    syncBodyModalState();
  };

  toggleCompareFullscreenButton.addEventListener("click", () => {
    setCompareFullscreenOpen(!compareFullscreenState.open);
  });

  closeCompareFullscreenButton?.addEventListener("click", (event) => {
    event.preventDefault();
    setCompareFullscreenOpen(false);
  });

  compareFullscreenModal.addEventListener("click", (event) => {
    const target = event.target;
    if (target instanceof HTMLElement && target.dataset.closeCompareFullscreen === "true") {
      setCompareFullscreenOpen(false);
    }
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && compareFullscreenState.open) {
      setCompareFullscreenOpen(false);
    }
  });

  window.addEventListener("resize", () => {
    if (compareFullscreenState.open) {
      syncCompareFullscreenWidth();
    }
  });

  setCompareFullscreenOpen(false);
}

function syncCompareFullscreenWidth() {
  if (!compareFullscreenCard) {
    return;
  }

  const reference = workspaceStage || compareBody?.closest(".workspace-stage") || document.body;
  const width = reference.getBoundingClientRect().width;
  if (!Number.isFinite(width) || width <= 0) {
    return;
  }

  const maxWidth = window.innerWidth - 40;
  compareFullscreenCard.style.width = `${Math.round(Math.min(width, maxWidth))}px`;
}

function initWorkspaceResizer() {
  if (!workspaceLayout || !workspaceResizer) {
    return;
  }

  const savedWidth = Number(window.localStorage.getItem("workspaceSidebarWidth") || "");
  if (Number.isFinite(savedWidth) && savedWidth > 0) {
    applyWorkspaceSidebarWidth(savedWidth);
  }

  workspaceResizer.addEventListener("pointerdown", (event) => {
    if (window.innerWidth <= 1180) {
      return;
    }

    resizeState.dragging = true;
    workspaceLayout.classList.add("resizing");
    workspaceResizer.setPointerCapture(event.pointerId);
    event.preventDefault();
  });

  workspaceResizer.addEventListener("pointermove", (event) => {
    if (!resizeState.dragging) {
      return;
    }

    const rect = workspaceLayout.getBoundingClientRect();
    applyWorkspaceSidebarWidth(rect.right - event.clientX);
  });

  const stopResizing = () => {
    if (!resizeState.dragging) {
      return;
    }

    resizeState.dragging = false;
    workspaceLayout.classList.remove("resizing");
  };

  workspaceResizer.addEventListener("pointerup", stopResizing);
  workspaceResizer.addEventListener("pointercancel", stopResizing);

  workspaceResizer.addEventListener("keydown", (event) => {
    if (window.innerWidth <= 1180) {
      return;
    }

    const currentWidth = getWorkspaceSidebarWidth();
    if (event.key === "ArrowLeft") {
      applyWorkspaceSidebarWidth(currentWidth - 24);
      event.preventDefault();
    }
    if (event.key === "ArrowRight") {
      applyWorkspaceSidebarWidth(currentWidth + 24);
      event.preventDefault();
    }
  });

  window.addEventListener("resize", () => {
    if (window.innerWidth <= 1180) {
      workspaceLayout.style.removeProperty("--sidebar-width");
      return;
    }

    applyWorkspaceSidebarWidth(getWorkspaceSidebarWidth(), false);
  });
}

function getWorkspaceSidebarWidth() {
  if (!workspaceLayout) {
    return 420;
  }

  const value = parseFloat(getComputedStyle(workspaceLayout).getPropertyValue("--sidebar-width"));
  return Number.isFinite(value) ? value : 420;
}

function applyWorkspaceSidebarWidth(nextWidth, persist = true) {
  if (!workspaceLayout || window.innerWidth <= 1180) {
    return;
  }

  const layoutWidth = workspaceLayout.getBoundingClientRect().width;
  const minWidth = 320;
  const maxWidth = Math.max(minWidth, Math.min(760, layoutWidth * 0.72));
  const clamped = Math.max(minWidth, Math.min(maxWidth, Number(nextWidth) || 420));

  workspaceLayout.style.setProperty("--sidebar-width", `${Math.round(clamped)}px`);
  if (persist) {
    window.localStorage.setItem("workspaceSidebarWidth", String(Math.round(clamped)));
  }
}

createForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(createForm);
  const button = createForm.querySelector("button");
  createResult.hidden = true;
  button.disabled = true;
  button.textContent = "上传中...";

  try {
    const response = await fetch("/api/collections/import-batch", {
      method: "POST",
      body: formData
    });
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "创建失败");
    }

    createResult.hidden = false;
    createResult.textContent = `批次已创建：${data.name}。已接收 ${data.fileCount} 个文件，导入 ${data.totalParsedCount} 条商品。`;
    createForm.reset();
    if (uploadEntryPanel) {
      uploadEntryPanel.hidden = true;
      uploadEntryPanel.setAttribute("aria-hidden", "true");
      uploadEntryPanel.classList.remove("is-open");
    }
    syncBodyModalState();
    toggleUploadEntryButton?.setAttribute("aria-expanded", "false");

    await loadCollections();
    await renderUploads();
    const created = state.collections.find((item) => item.id === data.id);
    if (created) {
      await selectCollection(created);
    }
  } catch (error) {
    createResult.hidden = false;
    createResult.textContent = error.message || "创建失败";
  } finally {
    button.disabled = false;
    button.textContent = "创建批次并上传";
  }
});

llmConfigForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(llmConfigForm);
  const payload = {
    provider: String(formData.get("provider") || "dashscope-compatible"),
    baseUrl: String(formData.get("baseUrl") || ""),
    model: String(formData.get("model") || ""),
    apiKey: String(formData.get("apiKey") || "")
  };

  try {
    const response = await fetch("/api/llm/config", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "保存配置失败");
    }

    llmConfigForm.elements.apiKey.value = "";
    applyAiConfig(data);
    renderAiStatus();
    renderLlmConfigStatus("大模型配置已保存。后续可按需使用这套模型做汇总分析和整理。");

    renderChatPanel();
    if (state.activeCollection) {
      resetProductsPanel();
    }
  } catch (error) {
    renderLlmConfigStatus(error.message || "保存配置失败");
  }
});

clearLlmConfigButton.addEventListener("click", async () => {
  try {
    const response = await fetch("/api/llm/config", {
      method: "DELETE"
    });
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "清空配置失败");
    }

    applyAiConfig(data);
    renderAiStatus();
    renderLlmConfigStatus("大模型配置已清空。");

    renderChatPanel();
    if (state.activeCollection) {
      resetProductsPanel();
    }
  } catch (error) {
    renderLlmConfigStatus(error.message || "清空配置失败");
  }
});

qjlLoginStartButton?.addEventListener("click", async () => {
  if (qjlState.loggedIn) {
    qjlState.homepageSwitcherOpen = !qjlState.homepageSwitcherOpen;
    renderQjlAuthState();
    return;
  }

  const idleText = qjlLoginStartButton.textContent;
  qjlLoginStartButton.disabled = true;
  qjlLoginStartButton.textContent = "准备中...";

  try {
    const response = await fetch("/api/qjl/auth/login/start", {
      method: "POST"
    });
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "获取群接龙登录链接失败");
    }

    applyQjlAccount(data);
    renderQjlAuthState("已获取登录链接。请完成群接龙登录后点击“检查登录”。");
    if (data.loginUrl) {
      window.open(data.loginUrl, "_blank", "noopener");
    }
  } catch (error) {
    renderQjlAuthState(error.message || "获取群接龙登录链接失败");
  } finally {
    qjlLoginStartButton.disabled = false;
    qjlLoginStartButton.textContent = idleText;
  }
});

qjlLoginCheckButton?.addEventListener("click", async () => {
  const idleText = qjlLoginCheckButton.textContent;
  qjlLoginCheckButton.disabled = true;
  qjlLoginCheckButton.textContent = "检查中...";

  try {
    const response = await fetch("/api/qjl/auth/poll", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({})
    });
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "检查群接龙登录失败");
    }

    applyQjlAccount(data);
    renderQjlAuthState(data.message || (data.loggedIn ? "群接龙已登录。" : "请先完成登录。"));
    if (state.activeCollection) {
      updateActiveCollectionCopy();
    }
  } catch (error) {
    renderQjlAuthState(error.message || "检查群接龙登录失败");
  } finally {
    qjlLoginCheckButton.disabled = false;
    qjlLoginCheckButton.textContent = idleText;
  }
});

qjlRefreshProfileButton?.addEventListener("click", async () => {
  const idleText = qjlRefreshProfileButton.textContent;
  qjlRefreshProfileButton.disabled = true;
  qjlRefreshProfileButton.textContent = "更新中...";

  try {
    const response = await fetch("/api/qjl/profile/refresh", {
      method: "POST"
    });
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "刷新用户画像失败");
    }

    applyQjlAccount(data);
    renderQjlAuthState(data.message || "用户画像已刷新。");
    if (state.activeCollection) {
      updateActiveCollectionCopy();
    }
  } catch (error) {
    renderQjlAuthState(error.message || "刷新用户画像失败");
  } finally {
    qjlRefreshProfileButton.disabled = false;
    qjlRefreshProfileButton.textContent = idleText;
  }
});

qjlLogoutButton?.addEventListener("click", async () => {
  const idleText = qjlLogoutButton.textContent;
  qjlLogoutButton.disabled = true;
  qjlLogoutButton.textContent = "退出中...";

  try {
    const response = await fetch("/api/qjl/account", {
      method: "DELETE"
    });
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "退出群接龙失败");
    }

    applyQjlAccount(data);
    renderQjlAuthState("已退出群接龙登录。");
    if (state.activeCollection) {
      updateActiveCollectionCopy();
    }
  } catch (error) {
    renderQjlAuthState(error.message || "退出群接龙失败");
  } finally {
    qjlLogoutButton.disabled = false;
    qjlLogoutButton.textContent = idleText;
  }
});

refreshUploads.addEventListener("click", async () => {
  await renderUploads();
  renderChatPanel();
});

if (sendSelectedToAiButton) {
  sendSelectedToAiButton.addEventListener("click", () => {
    const selectedIds = new Set(getSelectedUploadIds());
    const selectedUploads = (state.uploadBatches || []).filter((item) => selectedIds.has(item.id));

    if (!selectedUploads.length) {
      renderChatPanel("请先勾选至少一份货盘，再发送给 AI。");
      return;
    }

    chatInput.value = buildSelectedUploadsDraft(selectedUploads, chatInput.value);
    renderChatPanel(`已将 ${selectedUploads.length} 份货盘带入左侧对话框。你可以补充需求后再点击“发送”。`);
    chatInput.focus();
    const cursor = chatInput.value.length;
    chatInput.setSelectionRange(cursor, cursor);
    chatInput.scrollIntoView({ behavior: "smooth", block: "center" });
  });
}

compareForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(compareForm);
  const source = String(formData.get("source") || "local");

  if (source === "local" && !state.activeCollection) {
    return;
  }

  await renderComparison(state.activeCollection?.id || "", formData);
});

if (productForm) {
  productForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!state.activeCollection) {
      return;
    }

    await renderProducts(state.activeCollection.id, new FormData(productForm));
  });
}

clearChatButton.addEventListener("click", () => {
  state.chatMessages = [];
  renderChatPanel();
});

chatForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const message = String(chatInput.value || "").trim();

  if (!message) {
    renderChatPanel("先输入一个你想让 AI 回答的问题。");
    return;
  }

  chatInput.value = "";
  await submitUploadChat(message, {
    source: "chat",
    pendingButton: chatForm.querySelector("button[type='submit']"),
    pendingText: "分析中...",
    idleText: "发送"
  });
});

chatInput?.addEventListener("keydown", (event) => {
  if (event.key !== "Enter" || event.shiftKey || event.isComposing) {
    return;
  }

  event.preventDefault();
  chatForm.requestSubmit();
});

function buildSelectedUploadsDraft(selectedUploads, existingMessage = "") {
  const lines = selectedUploads.map(
    (item, index) => `${index + 1}. 批次：${item.batchName}；供应商：${item.supplierName}；货盘：${item.catalogName}`
  );

  const extraRequest = String(existingMessage || "").trim();
  return [
    "请基于以下已勾选货盘进行分析：",
    ...lines,
    "",
    extraRequest || "请补充你的分析需求："
  ].join("\n");
}

async function submitUploadChat(message, options = {}) {
  const selectedIds = getSelectedUploadIds();
  if (!selectedIds.length) {
    renderChatPanel("请先勾选至少一份货盘，再开始对话分析。");
    return;
  }

  if (!aiState.enabled) {
    renderChatPanel("当前还没有可用的大模型配置，请先在设置里完成配置。");
    return;
  }

  const history = getChatMessages();
  const nextMessages = [...history, { role: "user", content: message }];
  state.chatMessages = nextMessages;
  renderChatPanel("", true);

  const button = options.pendingButton || null;
  const idleText = options.idleText || button?.textContent || "";
  if (button) {
    button.disabled = true;
    button.textContent = options.pendingText || "处理中...";
  }

  try {
    const response = await fetch("/api/upload-chat", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        uploadIds: selectedIds,
        message,
        history
      })
    });
    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "对话分析失败");
    }

    state.chatMessages = [
      ...nextMessages,
      {
        role: "assistant",
        content: data.answer || "AI 没有返回可显示的内容。"
      }
    ];
    renderChatPanel();
  } catch (error) {
    state.chatMessages = [
      ...nextMessages,
      {
        role: "assistant",
        content: `分析失败：${error.message || "请稍后重试"}`
      }
    ];
    renderChatPanel();
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = idleText;
    }
  }
}

async function loadCollections() {
  const response = await fetch("/api/collections");
  const data = await response.json();
  state.collections = data.items || [];
  renderCollections();
}

async function loadAiStatus() {
  const response = await fetch("/api/llm/config");
  const data = await response.json();
  applyAiConfig(data);
  renderAiStatus();
  renderLlmConfigStatus();
}

async function loadQjlAccount() {
  const response = await fetch("/api/qjl/account");
  const data = await response.json();
  applyQjlAccount(data);
  renderQjlAuthState();
}

function renderAiStatus(message) {
  if (!aiStatus) {
    updateDetailAiStatus(aiState.enabled ? aiState.model : "未配置");
    return;
  }

  if (message) {
    aiStatus.textContent = message;
    updateDetailAiStatus(message);
    return;
  }

  if (aiState.enabled) {
    const text = `AI 已启用，当前模型：${aiState.model}，来源：${aiState.source === "saved" ? "软件配置" : "环境变量"}`;
    aiStatus.textContent = text;
    updateDetailAiStatus(aiState.model);
    return;
  }

  aiStatus.textContent = "当前还没有可用的大模型配置。请先在左侧保存模型配置，再使用 AI 整理货盘。";
  updateDetailAiStatus("未配置");
}

function renderLlmConfigStatus(message) {
  if (message) {
    llmConfigStatus.textContent = message;
    updateSettingsSummary();
    return;
  }

  if (!aiState.hasApiKey) {
    llmConfigStatus.textContent = "当前未保存 API Key。配置完成后，可按需调用这套模型做汇总分析和整理。";
    updateSettingsSummary();
    return;
  }

  llmConfigStatus.textContent = `当前使用 ${aiState.model}，接口 ${aiState.baseUrl}，Key：${aiState.apiKeyHint}。`;
  updateSettingsSummary();
}

function applyAiConfig(data) {
  aiState.enabled = Boolean(data.enabled);
  aiState.provider = data.provider || "dashscope-compatible";
  aiState.model = data.model || "qwen-plus-latest";
  aiState.baseUrl = data.baseUrl || "";
  aiState.source = data.source || "missing";
  aiState.hasApiKey = Boolean(data.hasApiKey);
  aiState.apiKeyHint = data.apiKeyHint || "";
  aiState.loaded = true;

  llmConfigForm.elements.provider.value = aiState.provider;
  llmConfigForm.elements.baseUrl.value = aiState.baseUrl;
  llmConfigForm.elements.model.value = aiState.model;
  llmConfigForm.elements.apiKey.placeholder = aiState.hasApiKey
    ? `当前已保存：${aiState.apiKeyHint}，留空则保留`
    : "请输入 API Key";
}

function updateSettingsSummary() {
  if (!settingsSummaryText) {
    return;
  }

  settingsSummaryText.textContent = aiState.hasApiKey
    ? `${aiState.model} · 已配置`
    : "未配置";
}

function applyQjlAccount(data) {
  qjlState.loggedIn = Boolean(data?.loggedIn);
  qjlState.pendingLogin = Boolean(data?.pendingLogin);
  qjlState.loginUrl = data?.loginUrl || "";
  qjlState.account = data?.account || null;
  qjlState.loaded = true;
  if (!qjlState.loggedIn) {
    qjlState.homepageSwitcherOpen = false;
  }
}

function renderQjlAuthState(message) {
  if (!qjlAuthStatus) {
    return;
  }

  if (message) {
    qjlAuthStatus.textContent = message;
  } else if (qjlState.loggedIn && qjlState.account) {
    const account = qjlState.account;
    qjlAuthStatus.textContent = account.profileSummary || `已登录 ${account.ghName || account.uid}`;
  } else if (qjlState.pendingLogin) {
    qjlAuthStatus.textContent = "已生成登录链接。完成群接龙登录后，点击“检查登录”即可同步状态。";
  } else {
    qjlAuthStatus.textContent = "登录后会基于群接龙主页信息生成用户画像，并在货盘分析时自动带入。";
  }

  if (qjlAuthSummary) {
    qjlAuthSummary.textContent = qjlState.loggedIn
      ? `${qjlState.account?.ghName || qjlState.account?.uid || "已登录"}`
      : qjlState.pendingLogin
        ? "待确认"
      : "未登录";
  }

  if (qjlLoginStartButton) {
    qjlLoginStartButton.textContent = qjlState.loggedIn ? "切换主页" : "开始登录";
  }

  if (qjlLoginLink) {
    if (qjlState.pendingLogin && qjlState.loginUrl) {
      qjlLoginLink.hidden = false;
      qjlLoginLink.innerHTML = `登录链接已生成：<a href="${escapeHtml(qjlState.loginUrl)}" target="_blank" rel="noreferrer">打开群接龙登录页</a>`;
    } else {
      qjlLoginLink.hidden = true;
      qjlLoginLink.textContent = "";
    }
  }

  if (qjlProfilePreview) {
    const account = qjlState.account;
    if (qjlState.loggedIn && account) {
      qjlProfilePreview.hidden = false;
      qjlProfileName.textContent = account.ghName || account.nickname || account.uid;
      qjlProfileMeta.textContent = [
        account.uid ? `uid ${account.uid}` : "",
        account.fansNum ? `粉丝 ${account.fansNum}` : "",
        account.orderNum ? `历史单量 ${account.orderNum}` : ""
      ]
        .filter(Boolean)
        .join(" · ");
      qjlProfileSummary.textContent = account.profileSummary || "已登录群接龙。";
      qjlProfileTags.innerHTML = (account.profile?.tags || [])
        .map((tag) => `<span class="tag">${escapeHtml(tag)}</span>`)
        .join("");
    } else {
      qjlProfilePreview.hidden = true;
      qjlProfileTags.innerHTML = "";
    }
  }

  renderQjlHomepageMenu(qjlState.account);

  if (qjlRefreshProfileButton) {
    qjlRefreshProfileButton.disabled = !qjlState.loggedIn;
  }

  if (qjlLogoutButton) {
    qjlLogoutButton.disabled = !qjlState.loggedIn;
  }
}

function renderQjlHomepageMenu(account) {
  if (!qjlHomepageMenu) {
    return;
  }

  if (!qjlState.loggedIn || !qjlState.homepageSwitcherOpen) {
    qjlHomepageMenu.hidden = true;
    return;
  }

  const homes = Array.isArray(account?.homepageList) ? account.homepageList : [];
  qjlHomepageMenu.hidden = false;

  if (!homes.length) {
    qjlHomepageMenu.innerHTML = `<div class="homepage-menu-empty">当前没有可切换的主页。</div>`;
    return;
  }

  qjlHomepageMenu.innerHTML = homes
    .map((home) => {
      const isCurrent = home.ghCode === account.ghId;
      const stats = [home.fansNum ? `粉丝 ${home.fansNum}` : "", home.orderNum ? `订单 ${home.orderNum}` : ""]
        .filter(Boolean)
        .join(" · ");
      return `
        <button
          type="button"
          class="homepage-menu-item ${isCurrent ? "active-homepage" : ""}"
          data-gh-code="${escapeHtml(home.ghCode)}"
          ${isCurrent ? "disabled" : ""}
        >
          <strong>${escapeHtml(home.ghName || home.ghCode)}${isCurrent ? "（当前）" : ""}</strong>
          <span>${escapeHtml(stats || home.ghCode)}</span>
        </button>
      `;
    })
    .join("");

  for (const button of qjlHomepageMenu.querySelectorAll("[data-gh-code]")) {
    button.addEventListener("click", async () => {
      const ghCode = String(button.dataset.ghCode || "").trim();
      if (!ghCode || ghCode === String(qjlState.account?.ghId || "")) {
        return;
      }

      const originalHtml = button.innerHTML;
      button.disabled = true;
      button.innerHTML = `<strong>切换中...</strong><span>正在更新该主页画像</span>`;

      try {
        const response = await fetch("/api/qjl/homepage/switch", {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({ ghCode })
        });
        const data = await response.json();

        if (!response.ok) {
          throw new Error(data.error || "切换群接龙主页失败");
        }

        applyQjlAccount(data);
        qjlState.homepageSwitcherOpen = false;
        renderQjlAuthState(data.message || "已切换群接龙主页。");
        if (state.activeCollection) {
          updateActiveCollectionCopy();
        }
      } catch (error) {
        button.disabled = false;
        button.innerHTML = originalHtml;
        renderQjlAuthState(error.message || "切换群接龙主页失败");
      }
    });
  }
}

function buildActiveCollectionDescription(collection) {
  const base =
    collection.description ||
    "勾选货盘后，可以直接和 AI 对话，继续分析这批货盘的价格、品类和风险点。";
  const profileSummary = qjlState.account?.profileSummary || "";

  if (qjlState.loggedIn && profileSummary) {
    return `${base} 当前会结合群接龙画像一起分析：${profileSummary}`;
  }

  return base;
}

function updateActiveCollectionCopy() {
  if (!state.activeCollection || !detailDescription) {
    return;
  }

  detailDescription.textContent = buildActiveCollectionDescription(state.activeCollection);
}

function renderCollections() {
  if (!collectionList) {
    return;
  }

  if (!state.collections.length) {
    collectionList.innerHTML = `
      <div class="empty-card">
        <h3>还没有批次</h3>
        <p>先在左侧新建收集批次，并一次上传多个供应商货盘文件。</p>
      </div>
    `;
    return;
  }

  collectionList.innerHTML = state.collections
    .map((item) => {
      const active = state.activeCollection?.id === item.id ? "active" : "";
      return `
        <button class="collection-item ${active}" data-id="${item.id}">
          <div>
            <strong>${escapeHtml(item.name)}</strong>
            <span>${escapeHtml(item.description || "暂无说明")}</span>
          </div>
          <div class="collection-stats">
            <span>${item.uploadCount} 次上传</span>
            <span>${item.productCount} 条商品</span>
          </div>
        </button>
      `;
    })
    .join("");

  for (const button of collectionList.querySelectorAll("[data-id]")) {
    button.addEventListener("click", async () => {
      const id = button.dataset.id;
      const found = state.collections.find((item) => item.id === id);
      if (found) {
        await selectCollection(found);
      }
    });
  }
}

async function selectCollection(collection) {
  state.activeCollection = collection;
  renderCollections();

  emptyState.hidden = true;
  detailContent.hidden = false;

  detailBadge.textContent = `当前批次 · ${formatDate(collection.createdAt)}`;
  detailName.textContent = collection.name;
  detailDescription.textContent = buildActiveCollectionDescription(collection);
  if (detailUploadCount) {
    detailUploadCount.textContent = collection.uploadCount || 0;
  }
  if (detailProductCount) {
    detailProductCount.textContent = collection.productCount || 0;
  }

  const url = new URL(window.location.href);
  url.searchParams.set("collectionId", collection.id);
  history.replaceState({}, "", url);

  await Promise.all([
    renderComparison(collection.id, new FormData(compareForm))
  ]);
  renderChatPanel();
  resetProductsPanel();
}

function updateDetailAiStatus(text) {
  if (detailAiStatus) {
    detailAiStatus.textContent = text || "未配置";
  }
}

async function renderUploads() {
  uploadTable.innerHTML = `<div class="loading">正在读取上传记录...</div>`;
  const response = await fetch("/api/upload-batches");
  const data = await response.json();
  const items = data.items || [];
  state.uploadBatches = items;

  const selectedIds = new Set(getSelectedUploadIds());
  const validIds = items.filter((item) => selectedIds.has(item.id)).map((item) => item.id);
  state.selectedUploadIds = validIds;
  renderUploadsTable();
}

async function renderComparison(collectionId, formData) {
  const keyword = String(formData.get("keyword") || "");
  const groupBy = String(formData.get("groupBy") || "smart");
  const source = String(formData.get("source") || "local");
  const matchMode = String(formData.get("matchMode") || "brand_style");
  compareResult.innerHTML = `<div class="loading">正在计算比价结果...</div>`;

  try {
    if (source !== "local") {
      if (!keyword.trim()) {
        compareResult.innerHTML = `
          <div class="empty-card slim">
            <h3>请输入搜索关键词</h3>
            <p>外部平台比价需要先输入商品关键词，例如“香蕉”“女装外套”或具体 SKU。</p>
          </div>
        `;
        return;
      }

      const params = new URLSearchParams({
        keyword,
        source,
        matchMode
      });
      const response = await fetch(`/api/compare/external-search?${params.toString()}`);
      const data = await response.json();

      if (!response.ok) {
        compareResult.innerHTML = `
          <div class="empty-card slim">
            <h3>外部比价查询失败</h3>
            <p>${escapeHtml(data.error || "请稍后重试")}</p>
          </div>
        `;
        return;
      }

      const items = Array.isArray(data.items) ? data.items : [];
      if (!items.length) {
        compareResult.innerHTML = `
          <div class="empty-card slim">
            <h3>没有找到匹配结果</h3>
            <p>${matchMode === "brand_style" ? "当前已按相同品牌和款式筛选。你可以切换到“全部结果”，或调整关键词后再试。" : "试试更换关键词，或者切到别的平台继续搜索。"}</p>
          </div>
        `;
        return;
      }

      compareResult.innerHTML = `
        <div class="compare-summary">
          ${escapeHtml(data.platformLabel || source)} ${matchMode === "brand_style" ? "按相同品牌和款式筛选后" : "共"}返回 <strong>${data.total || items.length}</strong> 条候选商品
        </div>
        <div class="compare-cards">
          ${items
            .map(
              (item) => `
                <article class="compare-card external-compare-card">
                  <div class="compare-card-top">
                    <div class="external-compare-main">
                      <div class="external-compare-thumb">
                        ${
                          item.imageUrl
                            ? `<img src="${escapeHtml(item.imageUrl)}" alt="${escapeHtml(item.title || "商品图")}" loading="lazy" referrerpolicy="no-referrer" />`
                            : `<div class="external-compare-thumb-fallback">暂无图片</div>`
                        }
                      </div>
                      <div class="external-compare-copy">
                        <h4>${escapeHtml(item.title || "未命名商品")}</h4>
                        <p>${escapeHtml(item.shopName || "未知店铺")} · ${escapeHtml(item.platformLabel || data.platformLabel || source)}</p>
                      </div>
                    </div>
                    <div class="price-badge">
                      <strong>${escapeHtml(item.priceText || "-")}</strong>
                      <span>${escapeHtml(item.salesText || "外部价格")}</span>
                    </div>
                  </div>
                  <div class="compare-meta">
                    <span>${escapeHtml(item.itemId || "无商品 ID")}</span>
                    <span>${escapeHtml(item.location || "地区未知")}</span>
                    <span>${escapeHtml(item.detailUrl ? "可跳转详情" : "无详情链接")}</span>
                  </div>
                  ${item.detailUrl ? `<a class="table-link external-link" href="${escapeHtml(item.detailUrl)}" target="_blank" rel="noreferrer">打开商品详情</a>` : ""}
                </article>
              `
            )
            .join("")}
        </div>
      `;
      return;
    }

    const params = new URLSearchParams({ keyword, groupBy });
    const response = await fetch(`/api/collections/${collectionId}/compare?${params.toString()}`);
    const data = await response.json();
    const items = data.items || [];

    if (!items.length) {
      compareResult.innerHTML = `
        <div class="empty-card slim">
          <h3>没有找到可比价的数据</h3>
          <p>试试换个关键词，或者先等供应商上传带价格的货盘。</p>
        </div>
      `;
      return;
    }

    compareResult.innerHTML = `
      <div class="compare-summary">
        共找到 <strong>${data.totalGroups}</strong> 个可比价商品分组
      </div>
      <div class="compare-cards">
        ${items
          .map(
            (item) => `
              <article class="compare-card">
                <div class="compare-card-top">
                  <div>
                    <h4>${escapeHtml(item.productName)}</h4>
                    <p>${escapeHtml(item.sku || "无 SKU")} · ${escapeHtml(item.spec || "未填写规格")}</p>
                  </div>
                  <div class="price-badge">
                    <strong>${formatMoney(item.minPrice)}</strong>
                    <span>最低价</span>
                  </div>
                </div>

                <div class="compare-meta">
                  <span>${item.offerCount} 家供应商</span>
                  <span>价差 ${formatMoney(item.priceSpread)}</span>
                  <span>${escapeHtml(item.unit || "单位未填")}</span>
                </div>

                <div class="offer-list">
                  ${item.offers
                    .map(
                      (offer, index) => `
                        <div class="offer-row ${index === 0 ? "best" : ""}">
                          <strong>${escapeHtml(offer.supplierName)}</strong>
                          <span>${formatMoney(offer.price)}</span>
                          <span>起订 ${offer.moq ?? "-"}</span>
                          <span>库存 ${offer.stock ?? "-"}</span>
                        </div>
                      `
                    )
                    .join("")}
                </div>
              </article>
            `
          )
          .join("")}
      </div>
    `;
  } catch (error) {
    compareResult.innerHTML = `
      <div class="empty-card slim">
        <h3>比价请求失败</h3>
        <p>${escapeHtml(error.message || "请求未完成，请稍后重试。")}</p>
      </div>
    `;
  }
}

function bindUploadSelectionEvents(items) {
  for (const checkbox of uploadTable.querySelectorAll(".upload-checkbox")) {
    checkbox.addEventListener("change", () => {
      const nextIds = Array.from(uploadTable.querySelectorAll(".upload-checkbox:checked")).map(
        (input) => input.dataset.uploadId
      );
      state.selectedUploadIds = nextIds;
      state.chatMessages = [];
      updateUploadSelectionUi(items);
      renderChatPanel();
    });
  }

  if (toggleAllUploads) {
    toggleAllUploads.onchange = () => {
      state.selectedUploadIds = toggleAllUploads.checked
        ? items.map((item) => item.id)
        : [];
      state.chatMessages = [];
      renderUploadsTable();
      renderChatPanel();
    };
  }
}

function renderUploadsTable() {
  const items = state.uploadBatches || [];

  if (!items.length) {
    uploadTable.innerHTML = `
      <div class="empty-card slim">
        <h3>还没有历史货盘</h3>
        <p>先在左侧创建收集批次并上传货盘，历史记录会按时间自动汇总到这里。</p>
      </div>
    `;
    updateUploadSelectionUi(items);
    return;
  }

  const selectedIds = new Set(getSelectedUploadIds());
  uploadTable.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>选择</th>
          <th>供应商</th>
          <th>货盘名称</th>
          <th>批次</th>
          <th>日期</th>
          <th>商品数量</th>
          <th>商品品类</th>
          <th>代表商品</th>
          <th>货盘摘要</th>
        </tr>
      </thead>
      <tbody>
        ${items
          .map(
            (item) => `
              <tr>
                <td>
                  <label class="checkbox-label">
                    <input type="checkbox" class="upload-checkbox" data-upload-id="${item.id}" ${
                      selectedIds.has(item.id) ? "checked" : ""
                    } />
                  </label>
                </td>
                <td>
                  <span class="truncate-text supplier-name-cell" title="${escapeHtml(item.supplierName)}">
                    ${escapeHtml(item.supplierName)}
                  </span>
                </td>
                <td>
                  <a class="table-link truncate-text catalog-link" href="${escapeHtml(item.fileUrl)}" target="_blank" rel="noreferrer" title="${escapeHtml(item.catalogName)}">
                    ${escapeHtml(item.catalogName)}
                  </a>
                </td>
                <td>
                  <span class="truncate-text batch-name-cell" title="${escapeHtml(item.batchName)}">
                    ${escapeHtml(item.batchName)}
                  </span>
                </td>
                <td>${formatDate(item.uploadedAt)}</td>
                <td>${item.productCount}</td>
                <td class="wrap-cell">${escapeHtml(item.categorySummary)}</td>
                <td class="wrap-cell">${escapeHtml(item.productNamesSummary || "未识别")}</td>
                <td class="wrap-cell">${escapeHtml(item.catalogOverview || "-")}</td>
              </tr>
            `
          )
          .join("")}
      </tbody>
    </table>
  `;

  bindUploadSelectionEvents(items);
  updateUploadSelectionUi(items);
}

function updateUploadSelectionUi(items) {
  const selectedIds = getSelectedUploadIds();
  const total = items.length;
  const selected = selectedIds.length;

  if (selectedUploadCount) {
    selectedUploadCount.textContent = `已选 ${selected} 份`;
  }

  if (toggleAllUploads) {
    toggleAllUploads.checked = total > 0 && selected === total;
    toggleAllUploads.indeterminate = selected > 0 && selected < total;
  }

  if (sendSelectedToAiButton) {
    sendSelectedToAiButton.disabled = selected === 0;
    sendSelectedToAiButton.textContent = selected > 0 ? `发送给AI（${selected}）` : "发送给AI";
  }
}

function getSelectedUploadIds() {
  return state.selectedUploadIds || [];
}

function getChatMessages() {
  return state.chatMessages || [];
}

function renderMarkdownMessage(content) {
  const source = String(content || "").replace(/\r\n/g, "\n");
  const lines = source.split("\n");
  const blocks = [];
  let index = 0;

  while (index < lines.length) {
    const line = lines[index];
    const trimmed = line.trim();

    if (!trimmed) {
      index += 1;
      continue;
    }

    if (trimmed.startsWith("```")) {
      const codeLines = [];
      index += 1;
      while (index < lines.length && !lines[index].trim().startsWith("```")) {
        codeLines.push(lines[index]);
        index += 1;
      }
      if (index < lines.length) {
        index += 1;
      }
      blocks.push(`<pre><code>${escapeHtml(codeLines.join("\n"))}</code></pre>`);
      continue;
    }

    if (isMarkdownTable(lines, index)) {
      const tableLines = [lines[index], lines[index + 1]];
      index += 2;
      while (index < lines.length && lines[index].includes("|") && lines[index].trim()) {
        tableLines.push(lines[index]);
        index += 1;
      }
      blocks.push(renderMarkdownTable(tableLines));
      continue;
    }

    if (/^#{1,3}\s+/.test(trimmed)) {
      const level = Math.min(3, trimmed.match(/^#+/)[0].length);
      const text = trimmed.replace(/^#{1,3}\s+/, "");
      blocks.push(`<h${level}>${renderInlineMarkdown(text)}</h${level}>`);
      index += 1;
      continue;
    }

    if (/^>\s?/.test(trimmed)) {
      const quoteLines = [];
      while (index < lines.length && /^>\s?/.test(lines[index].trim())) {
        quoteLines.push(lines[index].trim().replace(/^>\s?/, ""));
        index += 1;
      }
      blocks.push(`<blockquote>${quoteLines.map((item) => renderInlineMarkdown(item)).join("<br />")}</blockquote>`);
      continue;
    }

    if (/^(-|\*)\s+/.test(trimmed)) {
      const items = [];
      while (index < lines.length && /^(-|\*)\s+/.test(lines[index].trim())) {
        items.push(lines[index].trim().replace(/^(-|\*)\s+/, ""));
        index += 1;
      }
      blocks.push(`<ul>${items.map((item) => `<li>${renderInlineMarkdown(item)}</li>`).join("")}</ul>`);
      continue;
    }

    if (/^\d+\.\s+/.test(trimmed)) {
      const items = [];
      while (index < lines.length && /^\d+\.\s+/.test(lines[index].trim())) {
        items.push(lines[index].trim().replace(/^\d+\.\s+/, ""));
        index += 1;
      }
      blocks.push(`<ol>${items.map((item) => `<li>${renderInlineMarkdown(item)}</li>`).join("")}</ol>`);
      continue;
    }

    if (/^---+$/.test(trimmed) || /^\*\*\*+$/.test(trimmed)) {
      blocks.push("<hr />");
      index += 1;
      continue;
    }

    const paragraphLines = [];
    while (index < lines.length) {
      const candidate = lines[index];
      const candidateTrimmed = candidate.trim();
      if (
        !candidateTrimmed ||
        candidateTrimmed.startsWith("```") ||
        /^#{1,3}\s+/.test(candidateTrimmed) ||
        /^>\s?/.test(candidateTrimmed) ||
        /^(-|\*)\s+/.test(candidateTrimmed) ||
        /^\d+\.\s+/.test(candidateTrimmed) ||
        /^---+$/.test(candidateTrimmed) ||
        /^\*\*\*+$/.test(candidateTrimmed) ||
        isMarkdownTable(lines, index)
      ) {
        break;
      }
      paragraphLines.push(candidateTrimmed);
      index += 1;
    }
    blocks.push(`<p>${paragraphLines.map((item) => renderInlineMarkdown(item)).join("<br />")}</p>`);
  }

  return blocks.join("");
}

function isMarkdownTable(lines, index) {
  if (index + 1 >= lines.length) {
    return false;
  }

  const header = lines[index];
  const separator = lines[index + 1];
  if (!header.includes("|") || !separator.includes("|")) {
    return false;
  }

  return /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/.test(separator.trim());
}

function renderMarkdownTable(tableLines) {
  const rows = tableLines.map(splitMarkdownTableRow).filter((row) => row.length);
  if (rows.length < 2) {
    return `<p>${tableLines.map((line) => renderInlineMarkdown(line)).join("<br />")}</p>`;
  }

  const header = rows[0];
  const body = rows.slice(2);

  return `
    <table>
      <thead>
        <tr>${header.map((cell) => `<th>${renderInlineMarkdown(cell)}</th>`).join("")}</tr>
      </thead>
      <tbody>
        ${body
          .map(
            (row) => `<tr>${row.map((cell) => `<td>${renderInlineMarkdown(cell)}</td>`).join("")}</tr>`
          )
          .join("")}
      </tbody>
    </table>
  `;
}

function splitMarkdownTableRow(line) {
  const trimmed = String(line || "").trim().replace(/^\|/, "").replace(/\|$/, "");
  return trimmed.split("|").map((cell) => cell.trim());
}

function renderInlineMarkdown(text) {
  let html = escapeHtml(text);
  html = html.replace(/`([^`]+)`/g, "<code>$1</code>");
  html = html.replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g, '<a href="$2" target="_blank" rel="noreferrer">$1</a>');
  html = html.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
  html = html.replace(/\*([^*]+)\*/g, "<em>$1</em>");
  return html;
}

function renderChatPanel(statusMessage = "", pending = false) {
  const items = state.uploadBatches || [];
  const selectedIds = new Set(getSelectedUploadIds());
  const selectedUploads = items.filter((item) => selectedIds.has(item.id));
  const messages = getChatMessages();

  if (statusMessage) {
    chatSelectionStatus.hidden = false;
    chatSelectionStatus.textContent = statusMessage;
  } else {
    chatSelectionStatus.hidden = true;
    chatSelectionStatus.textContent = "";
  }

  if (!messages.length && !pending) {
    chatMessages.innerHTML = "";
    return;
  }

  const renderedMessages = [...messages];
  if (pending) {
    renderedMessages.push({
      role: "assistant",
      content: "正在结合你勾选的货盘整理分析，请稍候..."
    });
  }

  chatMessages.innerHTML = renderedMessages
    .map(
      (message) => `
        <article class="chat-bubble ${message.role === "user" ? "user" : "assistant"}">
          <div class="chat-role">${message.role === "user" ? "你" : "AI 分析"}</div>
          <div class="chat-content">${message.role === "assistant" ? renderMarkdownMessage(message.content) : escapeHtml(message.content).replace(/\n/g, "<br />")}</div>
        </article>
      `
    )
    .join("");

  chatMessages.scrollTop = chatMessages.scrollHeight;
}

async function renderProducts(collectionId, formData) {
  if (!productTable || !productSummary) {
    return;
  }

  const keyword = String(formData.get("keyword") || "");
  productSummary.hidden = true;
  productTable.innerHTML = `<div class="loading">正在读取 AI 整理结果...</div>`;

  if (!aiState.enabled) {
    renderAiStatus();
    productTable.innerHTML = `
      <div class="empty-card slim">
        <h3>当前未启用 AI</h3>
        <p>请先在左侧“大模型配置”里保存模型和 API Key，然后再执行 AI 整理。</p>
      </div>
    `;
    return;
  }

  const params = new URLSearchParams({ keyword });
  const response = await fetch(`/api/collections/${collectionId}/organized-products?${params.toString()}`);
  const data = await response.json();

  if (!response.ok) {
    renderAiStatus(data.error || "AI 整理失败");
    productTable.innerHTML = `
      <div class="empty-card slim">
        <h3>AI 整理失败</h3>
        <p>${escapeHtml(data.error || "请稍后重试")}</p>
      </div>
    `;
    return;
  }

  renderAiStatus(`AI 已启用，当前模型：${data.model || aiState.model}`);
  const items = data.items || [];
  const columns = data.columns || [];

  if (!items.length) {
    productTable.innerHTML = `
      <div class="empty-card slim">
        <h3>没有找到 AI 整理结果</h3>
        <p>试试换个关键词，或者先确认这个批次已经成功导入商品。</p>
      </div>
    `;
    return;
  }

  productSummary.hidden = false;
  productSummary.innerHTML = `
    ${escapeHtml(data.summary || "")}
    <br />
    当前显示 <strong>${data.visibleCount || items.length}</strong> / <strong>${data.total}</strong> 条。
    ${data.truncated ? "为控制成本，本次只整理前 80 条。" : ""}
  `;
  productTable.innerHTML = `
    <table class="product-table">
      <thead>
        <tr>
          <th>供应商</th>
          ${columns.map((column) => `<th>${escapeHtml(column.label)}</th>`).join("")}
          <th>上传时间</th>
        </tr>
      </thead>
      <tbody>
        ${items
          .map(
            (item) => `
              <tr>
                <td>${escapeHtml(item.supplierName || "-")}</td>
                ${columns
                  .map((column) => `<td class="wrap-cell">${escapeHtml(item.values?.[column.key] || "-")}</td>`)
                  .join("")}
                <td>${formatDate(item.uploadedAt)}</td>
              </tr>
            `
          )
          .join("")}
      </tbody>
    </table>
  `;
}

function resetProductsPanel() {
  if (!productTable || !productSummary) {
    return;
  }

  productSummary.hidden = true;
  productTable.innerHTML = `
    <div class="empty-card slim">
      <h3>按需启动 AI 整理</h3>
      <p>上传后系统默认只导入数据。如果你想看逐条字段整理结果，再点上方“开始整理”。</p>
    </div>
  `;
}

function formatDate(value) {
  if (!value) {
    return "-";
  }

  return new Intl.DateTimeFormat("zh-CN", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  }).format(new Date(value));
}

function formatMoney(value) {
  if (value === null || value === undefined || Number.isNaN(value)) {
    return "-";
  }

  return `¥${Number(value).toFixed(2)}`;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
