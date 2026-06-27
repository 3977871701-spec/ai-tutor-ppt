const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();

// 设置演示文稿属性
pptx.author = 'AI辅导员团队';
pptx.title = '学院AI辅导员项目汇报';
pptx.subject = '智慧校园服务建设方案';
pptx.company = '学院信息化办公室';

// 配色方案 - 专业商业风格
const COLORS = {
  primary: '1A365D',      // 深蓝
  secondary: '2B6CB0',     // 中蓝
  accent: 'ED8936',        // 橙色强调
  success: '38A169',       // 绿色
  warning: 'D69E2E',       // 黄色
  white: 'FFFFFF',
  lightGray: 'F7FAFC',
  mediumGray: 'E2E8F0',
  darkGray: '2D3748',
  textDark: '1A202C',
  textLight: '718096',
  textMedium: '4A5568'
};

// 创建阴影的工厂函数
const makeShadow = () => ({
  type: 'outer',
  color: '000000',
  blur: 4,
  offset: 2,
  angle: 135,
  opacity: 0.12
});

// ========== 第1页：封面 ==========
let slide1 = pptx.addSlide();
slide1.background = { color: COLORS.primary };

// 顶部装饰线
slide1.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0.6, w: 10, h: 0.04,
  fill: { color: COLORS.accent }
});

// 主标题
slide1.addText('学院AI辅导员系统', {
  x: 0.5, y: 1.6, w: 9, h: 1,
  fontSize: 48, bold: true, color: COLORS.white,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// 副标题
slide1.addText('项目汇报方案', {
  x: 0.5, y: 2.6, w: 9, h: 0.7,
  fontSize: 32, color: COLORS.accent,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// 底部装饰线
slide1.addShape(pptx.shapes.RECTANGLE, {
  x: 3, y: 3.5, w: 4, h: 0.03,
  fill: { color: COLORS.white }
});

// 关键词
slide1.addText('RAG知识库  |  微信公众号  |  智能问答  |  数字化转型', {
  x: 0.5, y: 4, w: 9, h: 0.5,
  fontSize: 16, color: '90CDF4',
  align: 'center', fontFace: 'Microsoft YaHei'
});

// 日期和汇报人
slide1.addText('2026年4月', {
  x: 0.5, y: 5, w: 9, h: 0.4,
  fontSize: 14, color: 'BEE3F8',
  align: 'center', fontFace: 'Microsoft YaHei'
});

// ========== 第2页：目录 ==========
let slide2 = pptx.addSlide();
slide2.background = { color: COLORS.lightGray };

slide2.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide2.addText('目  录', {
  x: 0.5, y: 0.2, w: 9, h: 0.5,
  fontSize: 28, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

const tocItems = [
  { num: '01', title: '项目背景与目标' },
  { num: '02', title: '系统架构设计' },
  { num: '03', title: '核心功能介绍' },
  { num: '04', title: '技术实现方案' },
  { num: '05', title: '运营保障方案' },
  { num: '06', title: '预算估算' },
  { num: '07', title: '实施计划' }
];

tocItems.forEach((item, i) => {
  const y = 1.2 + i * 0.58;
  
  // 序号圆圈
  slide2.addShape(pptx.shapes.OVAL, {
    x: 1.2, y: y, w: 0.45, h: 0.45,
    fill: { color: COLORS.secondary }
  });
  
  slide2.addText(item.num, {
    x: 1.2, y: y, w: 0.45, h: 0.45,
    fontSize: 14, bold: true, color: COLORS.white,
    align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
  });
  
  slide2.addText(item.title, {
    x: 1.85, y: y, w: 6, h: 0.45,
    fontSize: 18, color: COLORS.textDark,
    valign: 'middle', fontFace: 'Microsoft YaHei'
  });
  
  // 连接线
  slide2.addShape(pptx.shapes.LINE, {
    x: 4.5, y: y + 0.22, w: 3.5, h: 0,
    line: { color: COLORS.mediumGray, width: 1, dashType: 'dash' }
  });
});

// ========== 第3页：项目背景 ==========
let slide3 = pptx.addSlide();
slide3.background = { color: COLORS.lightGray };

slide3.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide3.addText('01  项目背景与目标', {
  x: 0.5, y: 0.2, w: 9, h: 0.5,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 痛点分析
slide3.addText('📋 现状痛点', {
  x: 0.5, y: 1.1, w: 4, h: 0.4,
  fontSize: 18, bold: true, color: COLORS.primary,
  fontFace: 'Microsoft YaHei'
});

slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 1.55, w: 4.3, h: 2.3,
  fill: { color: COLORS.white },
  shadow: makeShadow(), rectRadius: 0.08
});

slide3.addText([
  { text: '• 学生咨询需求分散', options: { breakLine: true } },
  { text: '   奖助学金、宿舍、请假、消三比...', options: { breakLine: true, color: COLORS.textLight, fontSize: 11 } },
  { text: '• 辅导员工作繁重', options: { breakLine: true } },
  { text: '   重复性咨询占用大量时间', options: { breakLine: true, color: COLORS.textLight, fontSize: 11 } },
  { text: '• 回复不及时、标准程度低', options: { breakLine: true } },
  { text: '• 学生难以快速获取准确信息', options: {} }
], {
  x: 0.7, y: 1.7, w: 3.9, h: 2.0,
  fontSize: 13, color: COLORS.textDark,
  fontFace: 'Microsoft YaHei', valign: 'top'
});

// 机遇分析
slide3.addText('🚀 发展机遇', {
  x: 5.2, y: 1.1, w: 4, h: 0.4,
  fontSize: 18, bold: true, color: COLORS.success,
  fontFace: 'Microsoft YaHei'
});

slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.2, y: 1.55, w: 4.3, h: 2.3,
  fill: { color: COLORS.white },
  shadow: makeShadow(), rectRadius: 0.08
});

slide3.addText([
  { text: '• 高校数字化转型加速推进', options: { breakLine: true } },
  { text: '• 学生已习惯通过微信获取服务', options: { breakLine: true } },
  { text: '• AI技术成熟度已达实用水平', options: { breakLine: true } },
  { text: '• RAG技术有效提升回答准确率', options: { breakLine: true } },
  { text: '• 微信公众号覆盖率高、触达快', options: {} }
], {
  x: 5.4, y: 1.7, w: 3.9, h: 2.0,
  fontSize: 13, color: COLORS.textDark,
  fontFace: 'Microsoft YaHei', valign: 'top'
});

// 目标
slide3.addText('🎯 项目目标', {
  x: 0.5, y: 4.0, w: 9, h: 0.4,
  fontSize: 18, bold: true, color: COLORS.primary,
  fontFace: 'Microsoft YaHei'
});

slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.45, w: 9, h: 0.9,
  fill: { color: COLORS.primary },
  shadow: makeShadow(), rectRadius: 0.08
});

slide3.addText('基于RAG知识库的智能问答机器人，7×24小时即时响应，\n减轻辅导员工作负担，提升学生满意度，实现服务标准化', {
  x: 0.7, y: 4.55, w: 8.6, h: 0.7,
  fontSize: 14, color: COLORS.white,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// ========== 第4页：系统架构 ==========
let slide4 = pptx.addSlide();
slide4.background = { color: COLORS.lightGray };

slide4.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide4.addText('02  系统架构设计', {
  x: 0.5, y: 0.2, w: 9, h: 0.5,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 三层架构
const layerY = [1.3, 2.8, 4.3];
const layerH = 1.3;
const layerW = 9;

// 接入层
slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: layerY[0], w: layerW, h: layerH,
  fill: { color: 'E6FFFA' },
  line: { color: '319795', width: 2 },
  rectRadius: 0.08
});

slide4.addText('📱 接入层', {
  x: 0.7, y: layerY[0] + 0.1, w: 1.5, h: 0.35,
  fontSize: 14, bold: true, color: '285E61',
  fontFace: 'Microsoft YaHei'
});

slide4.addText('微信公众号  |  Web管理后台  |  API接口', {
  x: 2.3, y: layerY[0] + 0.4, w: 7, h: 0.5,
  fontSize: 16, color: COLORS.textDark,
  fontFace: 'Microsoft YaHei', valign: 'middle'
});

// 业务层
slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: layerY[1], w: layerW, h: layerH,
  fill: { color: 'EBF8FF' },
  line: { color: '3182CE', width: 2 },
  rectRadius: 0.08
});

slide4.addText('⚙️ 业务层', {
  x: 0.7, y: layerY[1] + 0.1, w: 1.5, h: 0.35,
  fontSize: 14, bold: true, color: '2B6CB0',
  fontFace: 'Microsoft YaHei'
});

slide4.addText('意图识别  |  RAG检索  |  AI回答生成  |  转人工处理  |  知识库管理', {
  x: 2.3, y: layerY[1] + 0.4, w: 7, h: 0.5,
  fontSize: 16, color: COLORS.textDark,
  fontFace: 'Microsoft YaHei', valign: 'middle'
});

// 数据层
slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: layerY[2], w: layerW, h: layerH,
  fill: { color: 'FAF5FF' },
  line: { color: '805AD5', width: 2 },
  rectRadius: 0.08
});

slide4.addText('💾 数据层', {
  x: 0.7, y: layerY[2] + 0.1, w: 1.5, h: 0.35,
  fontSize: 14, bold: true, color: '553C9A',
  fontFace: 'Microsoft YaHei'
});

slide4.addText('ChromaDB向量库  |  文档存储  |  配置管理  |  缓存服务', {
  x: 2.3, y: layerY[2] + 0.4, w: 7, h: 0.5,
  fontSize: 16, color: COLORS.textDark,
  fontFace: 'Microsoft YaHei', valign: 'middle'
});

// ========== 第5页：工作流程 ==========
let slide5 = pptx.addSlide();
slide5.background = { color: COLORS.lightGray };

slide5.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide5.addText('02  系统工作流程', {
  x: 0.5, y: 0.2, w: 9, h: 0.5,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 流程步骤
const flowSteps = [
  { step: '1', title: '学生提问', desc: '通过微信公众号\n发送问题' },
  { step: '2', title: '意图识别', desc: '判断问题类型\n奖助学金/宿舍/消三比' },
  { step: '3', title: 'RAG检索', desc: '从知识库\n检索相关内容' },
  { step: '4', title: 'AI生成', desc: 'Grok-2 API\n生成回答' },
  { step: '5', title: '智能回复', desc: '返回答案或\n转接人工' }
];

const flowBoxW = 1.7;
const flowStartX = 0.5;

flowSteps.forEach((item, i) => {
  const x = flowStartX + i * 1.9;
  
  slide5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.3, w: flowBoxW, h: 2.2,
    fill: { color: COLORS.white },
    shadow: makeShadow(), rectRadius: 0.1
  });
  
  // 步骤编号
  slide5.addShape(pptx.shapes.OVAL, {
    x: x + 0.6, y: 1.45, w: 0.5, h: 0.5,
    fill: { color: COLORS.secondary }
  });
  
  slide5.addText(item.step, {
    x: x + 0.6, y: 1.45, w: 0.5, h: 0.5,
    fontSize: 18, bold: true, color: COLORS.white,
    align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
  });
  
  slide5.addText(item.title, {
    x: x, y: 2.05, w: flowBoxW, h: 0.4,
    fontSize: 14, bold: true, color: COLORS.primary,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
  
  slide5.addText(item.desc, {
    x: x + 0.1, y: 2.5, w: flowBoxW - 0.2, h: 0.9,
    fontSize: 11, color: COLORS.textLight,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
  
  // 箭头
  if (i < 4) {
    slide5.addText('→', {
      x: x + flowBoxW, y: 2.0, w: 0.3, h: 0.5,
      fontSize: 24, color: COLORS.secondary,
      align: 'center', valign: 'middle'
    });
  }
});

// 转人工流程
slide5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 3.8, w: 9, h: 1.4,
  fill: { color: 'FFF5F5' },
  line: { color: 'E53E3E', width: 1 },
  rectRadius: 0.08
});

slide5.addText('🔴 复杂问题转人工流程', {
  x: 0.7, y: 3.95, w: 8.6, h: 0.35,
  fontSize: 14, bold: true, color: 'C53030',
  fontFace: 'Microsoft YaHei'
});

slide5.addText('当AI判断为复杂问题时 → 自动推送对应老师联系方式 → 张老师(奖助学金)/李老师(宿舍)/王老师(消三比)', {
  x: 0.7, y: 4.4, w: 8.6, h: 0.6,
  fontSize: 13, color: COLORS.textDark,
  fontFace: 'Microsoft YaHei'
});

// ========== 第6页：核心功能 ==========
let slide6 = pptx.addSlide();
slide6.background = { color: COLORS.lightGray };

slide6.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide6.addText('03  核心功能介绍', {
  x: 0.5, y: 0.2, w: 9, h: 0.5,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

const coreFeatures = [
  { icon: '💬', title: '即时问答', desc: '学生输入问题，系统3秒内返回答案，支持模糊查询和语义理解' },
  { icon: '🔍', title: '意图识别', desc: '自动识别5类问题：奖助学金/宿舍/消三比/转人工/其他' },
  { icon: '🧠', title: 'RAG知识库', desc: '基于ChromaDB向量数据库，支持PDF/Word文档解析和语义检索' },
  { icon: '🤖', title: 'AI解答', desc: 'Grok-2 API驱动，结合知识库生成准确、专业的回答' },
  { icon: '📢', title: '公告推送', desc: '自动抓取学校官网公告，实时更新知识库内容' },
  { icon: '📝', title: '转人工服务', desc: '复杂问题无缝转接对应辅导员，确保服务闭环' }
];

coreFeatures.forEach((feat, i) => {
  const row = Math.floor(i / 2);
  const col = i % 2;
  const x = 0.5 + col * 4.7;
  const y = 1.1 + row * 1.4;
  
  slide6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: 4.4, h: 1.25,
    fill: { color: COLORS.white },
    shadow: makeShadow(), rectRadius: 0.08
  });
  
  // 图标
  slide6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.15, y: y + 0.25, w: 0.7, h: 0.7,
    fill: { color: COLORS.secondary },
    rectRadius: 0.1
  });
  
  slide6.addText(feat.icon, {
    x: x + 0.15, y: y + 0.25, w: 0.7, h: 0.7,
    fontSize: 24, align: 'center', valign: 'middle'
  });
  
  slide6.addText(feat.title, {
    x: x + 1.0, y: y + 0.2, w: 3.2, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.primary,
    fontFace: 'Microsoft YaHei'
  });
  
  slide6.addText(feat.desc, {
    x: x + 1.0, y: y + 0.6, w: 3.2, h: 0.55,
    fontSize: 11, color: COLORS.textMedium,
    fontFace: 'Microsoft YaHei'
  });
});

// ========== 第7页：知识库类别 ==========
let slide7 = pptx.addSlide();
slide7.background = { color: COLORS.lightGray };

slide7.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide7.addText('03  知识库内容', {
  x: 0.5, y: 0.2, w: 9, h: 0.5,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

const kbCategories = [
  { icon: '🏅', title: '奖助学金', items: '国家奖学金、助学金\n勤工俭学、助学贷款', color: 'FED7AA' },
  { icon: '🏠', title: '宿舍管理', items: '宿舍分配、调换申请\n卫生检查、违规处理', color: 'C6F6D5' },
  { icon: '📝', title: '请假制度', items: '请假类型、审批流程\n证明材料、补办规定', color: 'BEE3F8' },
  { icon: '🔥', title: '消三比', items: '学业预警、心理排查\n违纪处理等三比内容', color: 'FED7E2' },
  { icon: '📚', title: '教务相关', items: '选课、转专业、休学复学\n考试安排、成绩查询', color: 'E9D8FD' },
  { icon: '💼', title: '就业指导', items: '简历制作、面试技巧\n招聘信息、签约流程', color: 'FEEBC8' }
];

kbCategories.forEach((cat, i) => {
  const row = Math.floor(i / 3);
  const col = i % 3;
  const x = 0.4 + col * 3.15;
  const y = 1.1 + row * 2.1;
  
  slide7.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: 3.0, h: 1.9,
    fill: { color: COLORS.white },
    shadow: makeShadow(), rectRadius: 0.08
  });
  
  // 顶部色条
  slide7.addShape(pptx.shapes.RECTANGLE, {
    x: x, y: y, w: 3.0, h: 0.08,
    fill: { color: cat.color.replace(/^#/, '') ? COLORS.secondary : COLORS.secondary }
  });
  
  slide7.addText(`${cat.icon} ${cat.title}`, {
    x: x, y: y + 0.2, w: 3.0, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.primary,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
  
  slide7.addText(cat.items, {
    x: x + 0.15, y: y + 0.7, w: 2.7, h: 1.0,
    fontSize: 11, color: COLORS.textMedium,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
});

// ========== 第8页：技术方案 ==========
let slide8 = pptx.addSlide();
slide8.background = { color: COLORS.lightGray };

slide8.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide8.addText('04  技术实现方案', {
  x: 0.5, y: 0.2, w: 9, h: 0.5,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 技术栈表格
const techStack = [
  ['技术组件', '选型方案', '说明'],
  ['后端框架', 'FastAPI', '高性能Python异步框架'],
  ['向量数据库', 'ChromaDB', '轻量级向量数据库，支持语义检索'],
  ['嵌入模型', 'sentence-transformers', '文本向量化，支持中文'],
  ['AI模型', 'xAI Grok-2 API', '强大的对话生成能力'],
  ['微信公众号', '公众平台API', '消息接收与回复'],
  ['文档解析', 'PDF/Word Parser', '支持政策文件导入'],
  ['网页抓取', 'BeautifulSoup4', '学校公告自动抓取']
];

slide8.addTable(techStack, {
  x: 0.5, y: 1.1, w: 9, h: 3.2,
  colW: [2.2, 2.5, 4.3],
  border: { pt: 0.5, color: COLORS.mediumGray },
  fontFace: 'Microsoft YaHei',
  fontSize: 12,
  color: COLORS.textDark,
  align: 'left',
  valign: 'middle'
});

// 样式化表格
slide8.addShape(pptx.shapes.RECTANGLE, {
  x: 0.5, y: 1.1, w: 9, h: 0.45,
  fill: { color: COLORS.primary }
});

slide8.addText('技术组件                    选型方案                        说明', {
  x: 0.6, y: 1.15, w: 8.8, h: 0.35,
  fontSize: 12, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei'
});

// 技术优势
slide8.addText('✨ 技术优势', {
  x: 0.5, y: 4.5, w: 9, h: 0.35,
  fontSize: 16, bold: true, color: COLORS.primary,
  fontFace: 'Microsoft YaHei'
});

slide8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.9, w: 9, h: 0.7,
  fill: { color: COLORS.white },
  shadow: makeShadow(), rectRadius: 0.06
});

slide8.addText('跨平台支持（Windows/macOS/Linux）  |  轻量级部署  |  易于扩展  |  低运维成本', {
  x: 0.5, y: 4.9, w: 9, h: 0.7,
  fontSize: 14, color: COLORS.textMedium,
  align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
});

// ========== 第9页：运营方案 ==========
let slide9 = pptx.addSlide();
slide9.background = { color: COLORS.lightGray };

slide9.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide9.addText('05  运营保障方案', {
  x: 0.5, y: 0.2, w: 9, h: 0.5,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

const opsPlan = [
  { icon: '👥', title: '知识库运营', items: '• 专人负责知识库日常更新\n• 定期审核答案准确率\n• 收集学生反馈持续优化' },
  { icon: '📊', title: '数据监控', items: '• 实时监控服务可用性\n• 分析高频问题分布\n• 定期生成运营报告' },
  { icon: '🛡️', title: '内容安全', items: '• 内置敏感词过滤\n• 回答内容审核机制\n• 敏感问题自动转人工' },
  { icon: '🔧', title: '运维保障', items: '• 多副本部署容灾\n• 自动故障检测恢复\n• 7×24小时技术支持' }
];

opsPlan.forEach((item, i) => {
  const row = Math.floor(i / 2);
  const col = i % 2;
  const x = 0.5 + col * 4.7;
  const y = 1.1 + row * 2.1;
  
  slide9.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: 4.4, h: 1.9,
    fill: { color: COLORS.white },
    shadow: makeShadow(), rectRadius: 0.08
  });
  
  slide9.addText(item.icon, {
    x: x + 0.2, y: y + 0.15, w: 0.6, h: 0.6,
    fontSize: 28, align: 'center', valign: 'middle'
  });
  
  slide9.addText(item.title, {
    x: x + 0.85, y: y + 0.2, w: 3.3, h: 0.4,
    fontSize: 16, bold: true, color: COLORS.primary,
    fontFace: 'Microsoft YaHei'
  });
  
  slide9.addText(item.items, {
    x: x + 0.2, y: y + 0.7, w: 4.0, h: 1.1,
    fontSize: 11, color: COLORS.textMedium,
    fontFace: 'Microsoft YaHei'
  });
});

// ========== 第10页：预算估算 ==========
let slide10 = pptx.addSlide();
slide10.background = { color: COLORS.lightGray };

slide10.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide10.addText('06  预算估算', {
  x: 0.5, y: 0.2, w: 9, h: 0.5,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 预算方案
const budgetPlans = [
  { 
    type: '保守型', 
    amount: '¥30,900', 
    period: '季度',
    color: '48BB78',
    items: [
      '• API调用量：有限额度',
      '• 服务器：基础配置',
      '• 知识库规模：500条',
      '• 适用：小规模试点'
    ]
  },
  { 
    type: '均衡型', 
    amount: '¥64,500', 
    period: '季度',
    color: '4299E1',
    items: [
      '• API调用量：标准额度',
      '• 服务器：中等配置',
      '• 知识库规模：2000条',
      '• 适用：规模运营'
    ]
  },
  { 
    type: '激进型', 
    amount: '¥99,000', 
    period: '季度',
    color: 'ED8936',
    items: [
      '• API调用量：无限额度',
      '• 服务器：高可用配置',
      '• 知识库规模：无限扩展',
      '• 适用：全量上线'
    ]
  }
];

budgetPlans.forEach((plan, i) => {
  const x = 0.5 + i * 3.15;
  
  slide10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.1, w: 3.0, h: 3.5,
    fill: { color: COLORS.white },
    shadow: makeShadow(), rectRadius: 0.1
  });
  
  // 顶部标签
  slide10.addShape(pptx.shapes.RECTANGLE, {
    x: x, y: 1.1, w: 3.0, h: 0.5,
    fill: { color: plan.color }
  });
  
  slide10.addText(plan.type, {
    x: x, y: 1.15, w: 3.0, h: 0.4,
    fontSize: 16, bold: true, color: COLORS.white,
    align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
  });
  
  // 金额
  slide10.addText(plan.amount, {
    x: x, y: 1.75, w: 3.0, h: 0.6,
    fontSize: 26, bold: true, color: plan.color,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
  
  slide10.addText(plan.period, {
    x: x, y: 2.3, w: 3.0, h: 0.3,
    fontSize: 12, color: COLORS.textLight,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
  
  // 分隔线
  slide10.addShape(pptx.shapes.LINE, {
    x: x + 0.3, y: 2.7, w: 2.4, h: 0,
    line: { color: COLORS.mediumGray, width: 1 }
  });
  
  // 详细项
  slide10.addText(plan.items.join('\n'), {
    x: x + 0.2, y: 2.85, w: 2.6, h: 1.6,
    fontSize: 11, color: COLORS.textMedium,
    fontFace: 'Microsoft YaHei'
  });
});

// 建议
slide10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.8, w: 9, h: 0.8,
  fill: { color: 'EBF8FF' },
  line: { color: '3182CE', width: 1 },
  rectRadius: 0.06
});

slide10.addText('💡 建议：初期采用保守型方案（¥30,900/季度），3-6个月后根据运营数据切换至均衡型方案', {
  x: 0.7, y: 4.9, w: 8.6, h: 0.6,
  fontSize: 14, color: '2B6CB0',
  align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
});

// ========== 第11页：实施计划 ==========
let slide11 = pptx.addSlide();
slide11.background = { color: COLORS.lightGray };

slide11.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide11.addText('07  实施计划', {
  x: 0.5, y: 0.2, w: 9, h: 0.5,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

const phases = [
  { 
    phase: '第一阶段', 
    time: '第1-4周',
    title: '筹备与部署', 
    status: 'color: 48BB78',
    items: [
      '✓ 完成微信公众号认证',
      '✓ 服务器环境搭建',
      '✓ 系统部署与调试',
      '✓ 初始知识库建设'
    ]
  },
  { 
    phase: '第二阶段', 
    time: '第5-8周',
    title: '试点运营', 
    status: 'color: 4299E1',
    items: [
      '✓ 小范围试点运行',
      '✓ 收集用户反馈',
      '✓ 知识库持续优化',
      '✓ 系统稳定性监控'
    ]
  },
  { 
    phase: '第三阶段', 
    time: '第9-12周',
    title: '全面上线', 
    status: 'color: ED8936',
    items: [
      '✓ 全量学生开放使用',
      '✓ 运营数据分析',
      '✓ 功能迭代优化',
      '✓ 长期运维体系建立'
    ]
  }
];

phases.forEach((ph, i) => {
  const x = 0.5 + i * 3.15;
  
  slide11.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.1, w: 3.0, h: 3.5,
    fill: { color: COLORS.white },
    shadow: makeShadow(), rectRadius: 0.1
  });
  
  // 阶段标签
  slide11.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.2, y: 1.25, w: 2.6, h: 0.45,
    fill: { color: COLORS.secondary },
    rectRadius: 0.06
  });
  
  slide11.addText(ph.phase, {
    x: x + 0.2, y: 1.25, w: 2.6, h: 0.45,
    fontSize: 14, bold: true, color: COLORS.white,
    align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
  });
  
  // 时间
  slide11.addText(ph.time, {
    x: x, y: 1.85, w: 3.0, h: 0.4,
    fontSize: 18, bold: true, color: COLORS.primary,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
  
  // 标题
  slide11.addText(ph.title, {
    x: x, y: 2.3, w: 3.0, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.textDark,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
  
  // 分隔线
  slide11.addShape(pptx.shapes.LINE, {
    x: x + 0.3, y: 2.75, w: 2.4, h: 0,
    line: { color: COLORS.mediumGray, width: 1 }
  });
  
  // 详细项
  slide11.addText(ph.items.join('\n'), {
    x: x + 0.25, y: 2.9, w: 2.5, h: 1.6,
    fontSize: 11, color: COLORS.textMedium,
    fontFace: 'Microsoft YaHei'
  });
});

// 里程碑
slide11.addText('📍 关键里程碑', {
  x: 0.5, y: 4.75, w: 9, h: 0.35,
  fontSize: 14, bold: true, color: COLORS.primary,
  fontFace: 'Microsoft YaHei'
});

slide11.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 5.15, w: 9, h: 0.5,
  fill: { color: COLORS.white },
  shadow: makeShadow(), rectRadius: 0.06
});

slide11.addText('第4周完成内测  →  第8周试点反馈  →  第12周全量上线', {
  x: 0.5, y: 5.15, w: 9, h: 0.5,
  fontSize: 14, color: COLORS.textMedium,
  align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
});

// ========== 第12页：总结与展望 ==========
let slide12 = pptx.addSlide();
slide12.background = { color: COLORS.primary };

// 装饰线
slide12.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0.5, w: 10, h: 0.04,
  fill: { color: COLORS.accent }
});

slide12.addText('总结与展望', {
  x: 0.5, y: 0.8, w: 9, h: 0.7,
  fontSize: 32, bold: true, color: COLORS.white,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// 核心价值
const values = [
  { icon: '⚡', text: '提升效率', desc: '7×24小时即时响应' },
  { icon: '👥', text: '减轻负担', desc: '减少60%重复咨询' },
  { icon: '✅', text: '服务标准化', desc: '统一回答口径' },
  { icon: '📈', text: '持续优化', desc: '数据驱动迭代' }
];

values.forEach((val, i) => {
  const x = 0.5 + i * 2.35;
  
  slide12.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.8, w: 2.2, h: 1.6,
    fill: { color: 'FFFFFF', transparency: 15 },
    rectRadius: 0.1
  });
  
  slide12.addText(val.icon, {
    x: x, y: 1.95, w: 2.2, h: 0.5,
    fontSize: 28, align: 'center', valign: 'middle'
  });
  
  slide12.addText(val.text, {
    x: x, y: 2.5, w: 2.2, h: 0.4,
    fontSize: 16, bold: true, color: COLORS.white,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
  
  slide12.addText(val.desc, {
    x: x, y: 2.95, w: 2.2, h: 0.35,
    fontSize: 11, color: 'BEE3F8',
    align: 'center', fontFace: 'Microsoft YaHei'
  });
});

// 合作模式
slide12.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 3.7, w: 9, h: 1.0,
  fill: { color: '1A4073' },
  rectRadius: 0.08
});

slide12.addText('🤝 合作模式：技术支持 + 知识库共建 + 持续运营服务', {
  x: 0.5, y: 3.7, w: 9, h: 0.5,
  fontSize: 16, color: COLORS.white,
  align: 'center', valign: 'bottom', fontFace: 'Microsoft YaHei'
});

slide12.addText('共同打造智能化、人性化的校园服务新体验', {
  x: 0.5, y: 4.25, w: 9, h: 0.4,
  fontSize: 14, color: '90CDF4',
  align: 'center', fontFace: 'Microsoft YaHei'
});

// 底部
slide12.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 5.0, w: 10, h: 0.03,
  fill: { color: COLORS.accent }
});

slide12.addText('感谢各位领导的支持与指导', {
  x: 0.5, y: 5.15, w: 9, h: 0.4,
  fontSize: 16, color: 'BEE3F8',
  align: 'center', fontFace: 'Microsoft YaHei'
});

// ========== 第13页：附录-未解决项 ==========
let slide13 = pptx.addSlide();
slide13.background = { color: COLORS.lightGray };

slide13.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide13.addText('附录：待完成事项', {
  x: 0.5, y: 0.2, w: 9, h: 0.5,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 待完成项
slide13.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 1.1, w: 9, h: 2.0,
  fill: { color: 'FFF5F5' },
  line: { color: 'FC8181', width: 1 },
  rectRadius: 0.08
});

slide13.addText('⚠️ 待完成事项', {
  x: 0.7, y: 1.25, w: 8.6, h: 0.4,
  fontSize: 16, bold: true, color: 'C53030',
  fontFace: 'Microsoft YaHei'
});

slide13.addText([
  { text: '1. 微信公众号凭证配置', options: { bold: true, breakLine: true } },
  { text: '   AppID / AppSecret / Token / AESKey 待填写真实凭证', options: { color: COLORS.textLight, breakLine: true } },
  { text: '2. 服务部署启动', options: { bold: true, breakLine: true } },
  { text: '   端口8000待启动服务，配置公网访问', options: { color: COLORS.textLight, breakLine: true } },
  { text: '3. 知识库初始化', options: { bold: true, breakLine: true } },
  { text: '   补充完善各品类知识内容', options: { color: COLORS.textLight } }
], {
  x: 0.7, y: 1.7, w: 8.6, h: 1.3,
  fontSize: 13, color: COLORS.textDark,
  fontFace: 'Microsoft YaHei'
});

// 已完成项
slide13.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 3.3, w: 9, h: 2.2,
  fill: { color: 'F0FFF4' },
  line: { color: '68D391', width: 1 },
  rectRadius: 0.08
});

slide13.addText('✅ 已完成功能', {
  x: 0.7, y: 3.45, w: 8.6, h: 0.4,
  fontSize: 16, bold: true, color: '276749',
  fontFace: 'Microsoft YaHei'
});

slide13.addText([
  { text: '• FastAPI + ChromaDB RAG 知识库架构', options: { breakLine: true } },
  { text: '• 5类意图识别（奖助学金/宿舍/消三比/转人工/其他）', options: { breakLine: true } },
  { text: '• Grok-2 API AI解答功能', options: { breakLine: true } },
  { text: '• 转人工功能（张老师/李老师/王老师联系方式）', options: { breakLine: true } },
  { text: '• PDF/Word 文档解析', options: { breakLine: true } },
  { text: '• 学校公告自动抓取', options: { breakLine: true } },
  { text: '• 跨平台支持（Windows/macOS/Linux）', options: {} }
], {
  x: 0.7, y: 3.9, w: 8.6, h: 1.5,
  fontSize: 12, color: COLORS.textDark,
  fontFace: 'Microsoft YaHei'
});

// 保存文件
const outputPath = '学院AI辅导员项目汇报.pptx';
pptx.writeFile({ fileName: outputPath })
  .then(() => {
    console.log(`✅ PPT生成成功: ${outputPath}`);
    console.log(`📍 共13页幻灯片`);
    console.log(`📋 包含：项目背景、系统架构、功能介绍、技术方案、运营方案、预算估算、实施计划`);
  })
  .catch(err => {
    console.error('❌ 生成失败:', err);
  });
