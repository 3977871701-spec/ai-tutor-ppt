const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();

// 演示文稿属性
pptx.author = 'AI辅导员团队';
pptx.title = '学院AI辅导员系统介绍';
pptx.subject = '智能服务学生新方式';
pptx.company = 'AI辅导员项目组';

// 配色方案 - 商业提案风格，蓝白为主
const COLORS = {
  primary: '1A4B8C',      // 深蓝 - 主色
  secondary: '2563EB',   // 科技蓝
  accent: 'F59E0B',       // 琥珀强调
  white: 'FFFFFF',
  lightBg: 'F0F4F8',
  cardBg: 'FFFFFF',
  darkGray: '1E293B',
  mediumGray: '475569',
  lightGray: '94A3B8',
  border: 'CBD5E1',
  success: '059669',      // 绿色
  warning: 'D97706',      // 橙色
  danger: 'DC2626',       // 红色
  lightBlue: 'DBEAFE',
  lightAmber: 'FEF3C7'
};

// 通用阴影
const makeShadow = (opacity = 0.08) => ({
  type: 'outer',
  color: '000000',
  blur: 6,
  offset: 2,
  angle: 135,
  opacity
});

// =============================================
// 第1页：封面
// =============================================
let slide1 = pptx.addSlide();
slide1.background = { color: COLORS.primary };

// 顶部装饰条
slide1.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.08,
  fill: { color: COLORS.accent }
});

// 左上角装饰块
slide1.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0.08, w: 0.15, h: 2.5,
  fill: { color: COLORS.secondary, transparency: 40 }
});

// 主标题
slide1.addText('学院AI辅导员系统', {
  x: 0.5, y: 1.5, w: 9, h: 1.2,
  fontSize: 48, bold: true, color: COLORS.white,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// 副标题
slide1.addText('智能服务学生新方式', {
  x: 0.5, y: 2.7, w: 9, h: 0.7,
  fontSize: 28, color: COLORS.accent,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// 分隔线
slide1.addShape(pptx.shapes.RECTANGLE, {
  x: 3.5, y: 3.5, w: 3, h: 0.03,
  fill: { color: COLORS.white, transparency: 50 }
});

// 关键词标签
const keywords = ['RAG知识库', '智能问答', '微信公众号', '7×24服务'];
keywords.forEach((kw, i) => {
  const kwX = 1.5 + i * 1.9;
  slide1.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: kwX, y: 4.0, w: 1.7, h: 0.45,
    fill: { color: COLORS.white, transparency: 85 },
    line: { color: COLORS.white, width: 0.5 },
    rectRadius: 0.2
  });
  slide1.addText(kw, {
    x: kwX, y: 4.0, w: 1.7, h: 0.45,
    fontSize: 11, color: COLORS.white,
    align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
  });
});

// 底部信息
slide1.addText('客户版解决方案 · 2026', {
  x: 0.5, y: 5.1, w: 9, h: 0.4,
  fontSize: 14, color: '93C5FD',
  align: 'center', fontFace: 'Microsoft YaHei'
});

// =============================================
// 第2页：目录
// =============================================
let slide2 = pptx.addSlide();
slide2.background = { color: COLORS.lightBg };

// 左侧色块
slide2.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 3.2, h: 5.63,
  fill: { color: COLORS.primary }
});

slide2.addText('目录', {
  x: 0.3, y: 1.5, w: 2.6, h: 0.8,
  fontSize: 36, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei'
});

slide2.addText('CONTENTS', {
  x: 0.3, y: 2.3, w: 2.6, h: 0.4,
  fontSize: 12, color: '93C5FD',
  fontFace: 'Arial'
});

// 目录项
const tocItems = [
  { num: '01', title: '项目背景', sub: '学校学生服务现状与痛点' },
  { num: '02', title: '解决方案', sub: 'AI辅导员系统介绍' },
  { num: '03', title: '系统架构', sub: '技术架构与核心流程' },
  { num: '04', title: '核心功能', sub: '五大核心功能详解' },
  { num: '05', title: '知识库内容', sub: '学校办事指南与文件' },
  { num: '06', title: '运营保障', sub: '转人工机制与服务稳定性' },
  { num: '07', title: '效果预估', sub: '效率提升与成本节省' },
  { num: '08', title: '预算方案', sub: '三档方案详细对比' },
  { num: '09', title: '实施计划', sub: '三阶段里程碑' },
  { num: '10', title: '合作方式', sub: '落地模式与服务保障' },
];

tocItems.forEach((item, i) => {
  const row = Math.floor(i / 2);
  const col = i % 2;
  const x = 3.6 + col * 3.1;
  const y = 0.6 + row * 0.85;

  slide2.addText(item.num, {
    x: x, y: y, w: 0.5, h: 0.5,
    fontSize: 20, bold: true, color: COLORS.secondary,
    fontFace: 'Arial'
  });

  slide2.addText(item.title, {
    x: x + 0.55, y: y, w: 2.3, h: 0.35,
    fontSize: 14, bold: true, color: COLORS.darkGray,
    fontFace: 'Microsoft YaHei'
  });

  slide2.addText(item.sub, {
    x: x + 0.55, y: y + 0.35, w: 2.3, h: 0.3,
    fontSize: 10, color: COLORS.lightGray,
    fontFace: 'Microsoft YaHei'
  });
});

// =============================================
// 第3页：项目背景 - 痛点分析
// =============================================
let slide3 = pptx.addSlide();
slide3.background = { color: COLORS.lightBg };

// 顶部条
slide3.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide3.addText('01', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide3.addText('项目背景', {
  x: 1.0, y: 0.15, w: 3, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

slide3.addText('学校学生服务现状与痛点分析', {
  x: 1.0, y: 0.5, w: 5, h: 0.3,
  fontSize: 11, color: '93C5FD',
  fontFace: 'Microsoft YaHei', margin: 0
});

// 四大痛点卡片
const painPoints = [
  {
    icon: '😫',
    title: '咨询量大',
    desc: '辅导员每天需处理数百次重复性咨询，问题类型集中但耗时',
    stat: '人均100+/天'
  },
  {
    icon: '⏰',
    title: '响应不及时',
    desc: '学生等待回复时间长，尤其在非工作时间无法及时获取信息',
    stat: '平均等待4-8h'
  },
  {
    icon: '📋',
    title: '信息分散',
    desc: '奖助学金、宿舍、教务等政策文件分散在不同系统和网站',
    stat: '5+系统'
  },
  {
    icon: '😔',
    title: '满意度低',
    desc: '学生获取信息困难，辅导员工作压力大，双方体验均需改善',
    stat: '满意度<70%'
  }
];

painPoints.forEach((pp, i) => {
  const x = 0.4 + (i % 2) * 4.7;
  const y = 1.1 + Math.floor(i / 2) * 2.15;

  slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: 4.4, h: 2.0,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.1),
    rectRadius: 0.08
  });

  // 左侧色条
  slide3.addShape(pptx.shapes.RECTANGLE, {
    x: x, y: y + 0.3, w: 0.06, h: 1.4,
    fill: { color: COLORS.danger }
  });

  slide3.addText(pp.icon, {
    x: x + 0.2, y: y + 0.15, w: 0.6, h: 0.6,
    fontSize: 28
  });

  slide3.addText(pp.title, {
    x: x + 0.8, y: y + 0.2, w: 2.5, h: 0.4,
    fontSize: 16, bold: true, color: COLORS.darkGray,
    fontFace: 'Microsoft YaHei'
  });

  slide3.addText(pp.stat, {
    x: x + 3.2, y: y + 0.2, w: 1.0, h: 0.4,
    fontSize: 11, color: COLORS.danger, bold: true,
    fontFace: 'Microsoft YaHei', align: 'right'
  });

  slide3.addText(pp.desc, {
    x: x + 0.2, y: y + 0.7, w: 4.0, h: 1.1,
    fontSize: 12, color: COLORS.mediumGray,
    fontFace: 'Microsoft YaHei', valign: 'top'
  });
});

// =============================================
// 第4页：项目背景 - 行业趋势
// =============================================
let slide4 = pptx.addSlide();
slide4.background = { color: COLORS.lightBg };

slide4.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide4.addText('01', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide4.addText('行业趋势与政策背景', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 三大趋势
const trends = [
  {
    icon: '🏛️',
    title: '教育数字化战略',
    points: [
      '教育部《教育信息化2.0行动计划》全面实施',
      '智慧校园建设纳入高校评估核心指标',
      'AI赋能教育教学成为政策重点方向'
    ]
  },
  {
    icon: '📱',
    title: '学生行为变化',
    points: [
      '95后/00后学生完全习惯移动端获取服务',
      '微信日均使用时长超过4小时',
      '即时响应成为学生基本诉求'
    ]
  },
  {
    icon: '🤖',
    title: 'AI技术成熟',
    points: [
      '大语言模型推理能力达到专家水平',
      'RAG技术实现知识库精准检索',
      '落地成本大幅降低，ROI显著'
    ]
  }
];

trends.forEach((t, i) => {
  const x = 0.4 + i * 3.15;

  slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.1, w: 3.0, h: 3.8,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.1),
    rectRadius: 0.08
  });

  slide4.addShape(pptx.shapes.OVAL, {
    x: x + 1.1, y: 1.3, w: 0.8, h: 0.8,
    fill: { color: COLORS.secondary }
  });

  slide4.addText(t.icon, {
    x: x + 1.1, y: 1.35, w: 0.8, h: 0.7,
    fontSize: 28, align: 'center', valign: 'middle'
  });

  slide4.addText(t.title, {
    x: x + 0.15, y: 2.2, w: 2.7, h: 0.5,
    fontSize: 14, bold: true, color: COLORS.primary,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  t.points.forEach((p, j) => {
    slide4.addText('• ' + p, {
      x: x + 0.2, y: 2.8 + j * 0.7, w: 2.6, h: 0.65,
      fontSize: 11, color: COLORS.mediumGray,
      fontFace: 'Microsoft YaHei', valign: 'top'
    });
  });
});

// 底部结论
slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 5.05, w: 9.2, h: 0.45,
  fill: { color: COLORS.primary }
});

slide4.addText('✅ 结论：AI辅导员系统完美契合政策导向、技术趋势与学生需求，是智慧校园建设的最优切入点', {
  x: 0.5, y: 5.05, w: 9.0, h: 0.45,
  fontSize: 12, color: COLORS.white,
  fontFace: 'Microsoft YaHei', valign: 'middle'
});

// =============================================
// 第5页：项目背景 - 解决方案价值
// =============================================
let slide5 = pptx.addSlide();
slide5.background = { color: COLORS.lightBg };

slide5.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide5.addText('01', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide5.addText('解决方案核心价值', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 核心价值主张
slide5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 1.1, w: 9.2, h: 1.3,
  fill: { color: COLORS.primary },
  shadow: makeShadow(0.12),
  rectRadius: 0.1
});

slide5.addText('"7×24小时即时响应 · 知识库精准解答 · 复杂问题无缝转接"', {
  x: 0.5, y: 1.25, w: 9.0, h: 0.5,
  fontSize: 18, bold: true, color: COLORS.white,
  align: 'center', fontFace: 'Microsoft YaHei'
});

slide5.addText('基于RAG知识库的智能问答机器人，嵌入学校微信公众号，学生无需下载任何APP，对话即可获取学校文件与办事流程', {
  x: 0.5, y: 1.8, w: 9.0, h: 0.5,
  fontSize: 12, color: 'BFDBFE',
  align: 'center', fontFace: 'Microsoft YaHei'
});

// 四大价值
const values = [
  { icon: '⚡', title: '效率提升', desc: 'AI自动解答常见问题，辅导员专注于复杂任务', stat: '+80%' },
  { icon: '💰', title: '成本节省', desc: '减少重复性人力投入，降低服务运营成本', stat: '-60%' },
  { icon: '😊', title: '体验改善', desc: '学生随时获取即时服务，满意度显著提升', stat: '+95%' },
  { icon: '📊', title: '数据洞察', desc: '积累服务数据，为管理决策提供参考', stat: '可量化' }
];

values.forEach((v, i) => {
  const x = 0.4 + i * 2.35;

  slide5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 2.6, w: 2.2, h: 2.5,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.08),
    rectRadius: 0.08
  });

  slide5.addText(v.icon, {
    x: x, y: 2.75, w: 2.2, h: 0.6,
    fontSize: 30, align: 'center'
  });

  slide5.addText(v.title, {
    x: x, y: 3.4, w: 2.2, h: 0.4,
    fontSize: 14, bold: true, color: COLORS.darkGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide5.addText(v.stat, {
    x: x, y: 3.85, w: 2.2, h: 0.5,
    fontSize: 22, bold: true, color: COLORS.success,
    align: 'center', fontFace: 'Arial'
  });

  slide5.addText(v.desc, {
    x: x + 0.1, y: 4.35, w: 2.0, h: 0.7,
    fontSize: 10, color: COLORS.mediumGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
});

// =============================================
// 第6页：解决方案 - 产品定位
// =============================================
let slide6 = pptx.addSlide();
slide6.background = { color: COLORS.lightBg };

slide6.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide6.addText('02', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide6.addText('AI辅导员是什么', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 产品定义
slide6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 1.05, w: 9.2, h: 1.2,
  fill: { color: COLORS.lightBlue },
  rectRadius: 0.08
});

slide6.addText('产品定义', {
  x: 0.6, y: 1.15, w: 2, h: 0.35,
  fontSize: 12, bold: true, color: COLORS.secondary,
  fontFace: 'Microsoft YaHei'
});

slide6.addText('学院AI辅导员系统是一套基于RAG（检索增强生成）技术的智能问答解决方案，通过学校微信公众号这个学生日常使用最高频的入口，为学生提供7×24小时的校园事务智能问答服务。', {
  x: 0.6, y: 1.5, w: 8.8, h: 0.65,
  fontSize: 13, color: COLORS.darkGray,
  fontFace: 'Microsoft YaHei'
});

// 与传统方式对比
slide6.addText('与传统方式的对比', {
  x: 0.4, y: 2.4, w: 4, h: 0.4,
  fontSize: 14, bold: true, color: COLORS.darkGray,
  fontFace: 'Microsoft YaHei'
});

// 传统方式
slide6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 2.85, w: 4.4, h: 2.4,
  fill: { color: 'FEE2E2' },
  rectRadius: 0.08
});

slide6.addText('❌ 传统方式', {
  x: 0.6, y: 2.95, w: 4.0, h: 0.4,
  fontSize: 14, bold: true, color: COLORS.danger,
  fontFace: 'Microsoft YaHei'
});

const traditional = [
  '学生需记忆多个系统入口',
  '问题咨询需等待老师回复',
  '非工作时间无法获取服务',
  '信息更新不同步导致误导',
  '辅导员重复回答同一问题'
];

traditional.forEach((t, i) => {
  slide6.addText('• ' + t, {
    x: 0.6, y: 3.4 + i * 0.38, w: 4.0, h: 0.35,
    fontSize: 11, color: COLORS.mediumGray,
    fontFace: 'Microsoft YaHei'
  });
});

// AI方式
slide6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.2, y: 2.85, w: 4.4, h: 2.4,
  fill: { color: 'D1FAE5' },
  rectRadius: 0.08
});

slide6.addText('✅ AI辅导员', {
  x: 5.4, y: 2.95, w: 4.0, h: 0.4,
  fontSize: 14, bold: true, color: COLORS.success,
  fontFace: 'Microsoft YaHei'
});

const aiWays = [
  '微信公众号一键触达',
  '3秒内即时智能响应',
  '7×24全天候不间断服务',
  '知识库集中管理实时更新',
  '同一问题只回答一次'
];

aiWays.forEach((a, i) => {
  slide6.addText('• ' + a, {
    x: 5.4, y: 3.4 + i * 0.38, w: 4.0, h: 0.35,
    fontSize: 11, color: COLORS.mediumGray,
    fontFace: 'Microsoft YaHei'
  });
});

// =============================================
// 第7页：解决方案 - 功能全景
// =============================================
let slide7 = pptx.addSlide();
slide7.background = { color: COLORS.lightBg };

slide7.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide7.addText('02', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide7.addText('AI辅导员能做什么', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 功能矩阵
const functions = [
  {
    icon: '🎓',
    title: '奖助学金咨询',
    desc: '国家奖学金、助学金、勤工俭学、助学贷款等各类奖助政策的申请条件、流程、材料清单',
    tags: ['奖学金申请', '助学金', '助学贷款', '贫困生认定']
  },
  {
    icon: '🏠',
    title: '宿舍事务办理',
    desc: '调换宿舍、退宿办理、住宿申请、宿舍费用查询等宿舍相关全流程指引',
    tags: ['调换宿舍', '退宿', '住宿申请', '宿舍费用']
  },
  {
    icon: '✅',
    title: '消三比达标',
    desc: '消三比达标标准查询、预警消除申请流程、学业/综素/宿舍卫生各维度说明',
    tags: ['消三比', '学业预警', '达标申请', '预警消除']
  },
  {
    icon: '📄',
    title: '其他校园事务',
    desc: '请假流程、户籍办理、成绩证明、毕业手续等各类常见校园事务咨询',
    tags: ['请假', '成绩单', '户籍', '毕业']
  },
  {
    icon: '💬',
    title: '智能闲聊互动',
    desc: '友好的开场白和寒暄响应，提供有温度的服务体验，增强学生使用意愿',
    tags: ['问候', '闲聊', '引导']
  },
  {
    icon: '📞',
    title: '一键转人工',
    desc: '复杂问题或AI无法解答时，自动转接对应部门老师，提供姓名、电话、部门信息',
    tags: ['转人工', '联系老师', '工单创建']
  }
];

functions.forEach((f, i) => {
  const row = Math.floor(i / 2);
  const col = i % 2;
  const x = 0.4 + col * 4.7;
  const y = 1.05 + row * 1.5;

  slide7.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: 4.5, h: 1.35,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.06),
    rectRadius: 0.06
  });

  slide7.addText(f.icon + ' ' + f.title, {
    x: x + 0.15, y: y + 0.1, w: 4.2, h: 0.35,
    fontSize: 13, bold: true, color: COLORS.primary,
    fontFace: 'Microsoft YaHei'
  });

  slide7.addText(f.desc, {
    x: x + 0.15, y: y + 0.45, w: 4.2, h: 0.55,
    fontSize: 10, color: COLORS.mediumGray,
    fontFace: 'Microsoft YaHei', valign: 'top'
  });

  // Tags
  f.tags.slice(0, 2).forEach((tag, j) => {
    slide7.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.15 + j * 1.3, y: y + 1.05, w: 1.2, h: 0.22,
      fill: { color: COLORS.lightBlue },
      rectRadius: 0.1
    });
    slide7.addText(tag, {
      x: x + 0.15 + j * 1.3, y: y + 1.05, w: 1.2, h: 0.22,
      fontSize: 8, color: COLORS.secondary,
      align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
    });
  });
});

// =============================================
// 第8页：系统架构
// =============================================
let slide8 = pptx.addSlide();
slide8.background = { color: COLORS.lightBg };

slide8.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide8.addText('03', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide8.addText('系统架构', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 技术架构图 - 四层
const layers = [
  { name: '接入层', color: '059669', items: ['微信公众号', '网页端', '小程序'] },
  { name: '业务层', color: '2563EB', items: ['意图识别', 'RAG检索', 'AI回复生成', '转人工处理'] },
  { name: '数据层', color: '7C3AED', items: ['ChromaDB知识库', '向量嵌入模型', '学校文档库'] },
  { name: '支撑层', color: 'DC2626', items: ['FastAPI服务', 'Grok-2 API', '学校公告抓取'] }
];

layers.forEach((layer, i) => {
  const y = 1.1 + i * 1.05;
  const layerW = 3.5;

  // 层级标签
  slide8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: y, w: 1.2, h: 0.85,
    fill: { color: layer.color },
    rectRadius: 0.06
  });

  slide8.addText(layer.name, {
    x: 0.4, y: y, w: 1.2, h: 0.85,
    fontSize: 13, bold: true, color: COLORS.white,
    align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
  });

  // 组件
  layer.items.forEach((item, j) => {
    const itemX = 1.8 + j * 1.65;
    slide8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: itemX, y: y, w: 1.55, h: 0.85,
      fill: { color: COLORS.cardBg },
      line: { color: layer.color, width: 1.5 },
      rectRadius: 0.06
    });

    slide8.addText(item, {
      x: itemX, y: y, w: 1.55, h: 0.85,
      fontSize: 10, color: COLORS.darkGray,
      align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
    });
  });
});

// 数据流向箭头说明
slide8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 5.05, w: 9.2, h: 0.45,
  fill: { color: COLORS.cardBg },
  line: { color: COLORS.border, width: 1 },
  rectRadius: 0.05
});

slide8.addText('📍 数据流：学生消息 → 微信公众号 → FastAPI → 意图识别 → RAG检索 → LLM生成 → 微信回复', {
  x: 0.5, y: 5.05, w: 9.0, h: 0.45,
  fontSize: 11, color: COLORS.mediumGray,
  fontFace: 'Microsoft YaHei', valign: 'middle'
});

// =============================================
// 第9页：系统架构 - 工作流程
// =============================================
let slide9 = pptx.addSlide();
slide9.background = { color: COLORS.lightBg };

slide9.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide9.addText('03', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide9.addText('核心工作流程', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 流程步骤
const flowSteps = [
  { step: '01', icon: '💬', title: '学生提问', desc: '学生通过微信公众号发送问题' },
  { step: '02', icon: '🔍', title: '意图识别', desc: 'AI识别5大类意图：奖助学金/宿舍/消三比/其他/闲聊' },
  { step: '03', icon: '📚', title: '知识检索', desc: 'ChromaDB向量检索，匹配最相关的知识文档' },
  { step: '04', icon: '🤖', title: 'AI生成', desc: 'Grok-2大模型结合知识库生成专业回答' },
  { step: '05', icon: '✅', title: '智能回复', desc: '3秒内通过微信推送精准答案给学生' }
];

flowSteps.forEach((s, i) => {
  const x = 0.3 + i * 1.92;

  // 步骤卡片
  slide9.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.2, w: 1.8, h: 2.6,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.08),
    rectRadius: 0.08
  });

  // 步骤编号圆
  slide9.addShape(pptx.shapes.OVAL, {
    x: x + 0.6, y: 1.35, w: 0.6, h: 0.6,
    fill: { color: COLORS.secondary }
  });

  slide9.addText(s.step, {
    x: x + 0.6, y: 1.35, w: 0.6, h: 0.6,
    fontSize: 14, bold: true, color: COLORS.white,
    align: 'center', valign: 'middle', fontFace: 'Arial'
  });

  slide9.addText(s.icon, {
    x: x, y: 2.05, w: 1.8, h: 0.5,
    fontSize: 24, align: 'center'
  });

  slide9.addText(s.title, {
    x: x + 0.1, y: 2.55, w: 1.6, h: 0.4,
    fontSize: 12, bold: true, color: COLORS.darkGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide9.addText(s.desc, {
    x: x + 0.1, y: 2.95, w: 1.6, h: 0.75,
    fontSize: 9, color: COLORS.mediumGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  // 箭头
  if (i < 4) {
    slide9.addText('→', {
      x: x + 1.75, y: 2.1, w: 0.25, h: 0.5,
      fontSize: 18, color: COLORS.lightGray,
      align: 'center', valign: 'middle'
    });
  }
});

// 转人工流程
slide9.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 4.0, w: 9.2, h: 1.45,
  fill: { color: 'FFF7ED' },
  line: { color: COLORS.accent, width: 1.5 },
  rectRadius: 0.08
});

slide9.addText('🔀 转人工流程（复杂问题兜底）', {
  x: 0.6, y: 4.1, w: 8.8, h: 0.4,
  fontSize: 13, bold: true, color: COLORS.warning,
  fontFace: 'Microsoft YaHei'
});

slide9.addText('当学生请求转人工 / AI置信度低 / 系统判定为复杂问题时 → 自动推送对应部门老师联系方式（姓名+电话+负责领域）→ 学生可直接致电或前往办理 → 全程服务闭环', {
  x: 0.6, y: 4.5, w: 8.8, h: 0.85,
  fontSize: 11, color: COLORS.mediumGray,
  fontFace: 'Microsoft YaHei'
});

// =============================================
// 第10页：核心功能 - 意图识别
// =============================================
let slide10 = pptx.addSlide();
slide10.background = { color: COLORS.lightBg };

slide10.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide10.addText('04', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide10.addText('核心功能详解', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 功能1: 意图识别
slide10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 1.0, w: 9.2, h: 0.5,
  fill: { color: COLORS.secondary },
  rectRadius: 0.05
});

slide10.addText('🎯 功能一：意图识别引擎', {
  x: 0.5, y: 1.0, w: 9.0, h: 0.5,
  fontSize: 16, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', valign: 'middle'
});

// 五大意图
const intents = [
  { type: 'scholarship', icon: '🎓', name: '奖助学金', keywords: '奖学金,助学金,助学贷款,贫困生,补贴', desc: '国家助学金/奖学金/励志奖学金等申请条件、流程、材料' },
  { type: 'dormitory', icon: '🏠', name: '宿舍相关', keywords: '宿舍,床位,住宿,调换,退宿,退房,寝室', desc: '调换宿舍、退宿办理、住宿申请等流程指引' },
  { type: 'three_ratio', icon: '✅', name: '消三比', keywords: '消三比,三比,达标,预警,学分,绩点', desc: '消三比达标标准、预警消除申请流程说明' },
  { type: 'other', icon: '📄', name: '其他办事', keywords: '请假,成绩,证明,毕业,户籍', desc: '请假制度、成绩证明、毕业手续等咨询' },
  { type: 'greeting', icon: '💬', name: '闲聊/其他', keywords: '你好,您好,在吗,hi,hello', desc: '问候语响应及友好交互' }
];

intents.forEach((intent, i) => {
  const x = 0.4 + i * 1.84;

  slide10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.65, w: 1.74, h: 2.9,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.06),
    rectRadius: 0.06
  });

  slide10.addShape(pptx.shapes.OVAL, {
    x: x + 0.57, y: 1.8, w: 0.6, h: 0.6,
    fill: { color: COLORS.secondary }
  });

  slide10.addText(intent.icon, {
    x: x + 0.57, y: 1.82, w: 0.6, h: 0.55,
    fontSize: 22, align: 'center', valign: 'middle'
  });

  slide10.addText(intent.name, {
    x: x, y: 2.5, w: 1.74, h: 0.35,
    fontSize: 11, bold: true, color: COLORS.darkGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide10.addText('关键词:', {
    x: x + 0.1, y: 2.9, w: 1.54, h: 0.25,
    fontSize: 8, bold: true, color: COLORS.lightGray,
    fontFace: 'Microsoft YaHei'
  });

  slide10.addText(intent.keywords, {
    x: x + 0.1, y: 3.1, w: 1.54, h: 0.6,
    fontSize: 8, color: COLORS.mediumGray,
    fontFace: 'Microsoft YaHei'
  });

  slide10.addText(intent.desc, {
    x: x + 0.1, y: 3.7, w: 1.54, h: 0.75,
    fontSize: 8, color: COLORS.mediumGray,
    fontFace: 'Microsoft YaHei'
  });
});

// 技术说明
slide10.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 4.7, w: 9.2, h: 0.8,
  fill: { color: COLORS.lightBlue },
  rectRadius: 0.05
});

slide10.addText('技术实现：基于关键词匹配 + 语义理解的双层识别机制，支持同义词扩展和模糊匹配，识别准确率≥95%', {
  x: 0.5, y: 4.7, w: 9.0, h: 0.8,
  fontSize: 12, color: COLORS.secondary,
  fontFace: 'Microsoft YaHei', valign: 'middle'
});

// =============================================
// 第11页：核心功能 - RAG知识库
// =============================================
let slide11 = pptx.addSlide();
slide11.background = { color: COLORS.lightBg };

slide11.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide11.addText('04', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide11.addText('核心功能详解', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

slide11.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 1.0, w: 9.2, h: 0.5,
  fill: { color: COLORS.secondary },
  rectRadius: 0.05
});

slide11.addText('🧠 功能二：RAG检索增强知识库', {
  x: 0.5, y: 1.0, w: 9.0, h: 0.5,
  fontSize: 16, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', valign: 'middle'
});

// RAG流程
const ragSteps = [
  { icon: '📄', title: '文档解析', desc: '支持PDF/Word/TXT等格式，自动提取文本内容' },
  { icon: '✂️', title: '文本分块', desc: '智能分chunk，保留语义完整性，便于精准检索' },
  { icon: '🔢', title: '向量嵌入', desc: '多语言模型paraphrase-multilingual-MiniLM-L12-v2' },
  { icon: '💾', title: 'ChromaDB存储', desc: '向量数据库持久化存储，支持高速相似度检索' },
  { icon: '🔍', title: '语义检索', desc: '学生问题转向量，最匹配Top-K文档召回' },
  { icon: '🤖', title: 'LLM生成', desc: 'Grok-2结合检索内容生成准确、专业的回答' }
];

ragSteps.forEach((r, i) => {
  const row = Math.floor(i / 3);
  const col = i % 3;
  const x = 0.4 + col * 3.1;
  const y = 1.65 + row * 1.55;

  slide11.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: 2.95, h: 1.4,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.06),
    rectRadius: 0.06
  });

  slide11.addText(r.icon, {
    x: x + 0.1, y: y + 0.1, w: 0.5, h: 0.5,
    fontSize: 22
  });

  slide11.addText(r.title, {
    x: x + 0.6, y: y + 0.15, w: 2.2, h: 0.35,
    fontSize: 13, bold: true, color: COLORS.darkGray,
    fontFace: 'Microsoft YaHei'
  });

  slide11.addText(r.desc, {
    x: x + 0.1, y: y + 0.6, w: 2.75, h: 0.7,
    fontSize: 10, color: COLORS.mediumGray,
    fontFace: 'Microsoft YaHei'
  });
});

// 知识库优势
slide11.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 4.85, w: 9.2, h: 0.65,
  fill: { color: 'F0FDF4' },
  line: { color: COLORS.success, width: 1 },
  rectRadius: 0.05
});

slide11.addText('✅ 优势：答案基于学校官方文件，回答准确可信；支持增量更新，文件修改后一键同步；跨平台部署，Windows/Mac/Linux均可运行', {
  x: 0.5, y: 4.85, w: 9.0, h: 0.65,
  fontSize: 11, color: COLORS.success,
  fontFace: 'Microsoft YaHei', valign: 'middle'
});

// =============================================
// 第12页：核心功能 - AI智能解答
// =============================================
let slide12 = pptx.addSlide();
slide12.background = { color: COLORS.lightBg };

slide12.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide12.addText('04', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide12.addText('核心功能详解', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

slide12.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 1.0, w: 9.2, h: 0.5,
  fill: { color: COLORS.secondary },
  rectRadius: 0.05
});

slide12.addText('🤖 功能三：AI智能解答（Grok-2大模型）', {
  x: 0.5, y: 1.0, w: 9.0, h: 0.5,
  fontSize: 16, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', valign: 'middle'
});

// 模型能力
slide12.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 1.65, w: 5.8, h: 2.5,
  fill: { color: COLORS.cardBg },
  shadow: makeShadow(0.06),
  rectRadius: 0.08
});

slide12.addText('Grok-2 模型能力', {
  x: 0.6, y: 1.75, w: 5.4, h: 0.4,
  fontSize: 14, bold: true, color: COLORS.primary,
  fontFace: 'Microsoft YaHei'
});

const modelCaps = [
  '✓ 强大的中文语义理解能力',
  '✓ 专业领域知识问答专家',
  '✓ 支持多轮对话上下文记忆',
  '✓ 生成内容结构清晰、表达专业',
  '✓ 可处理复杂模糊问题',
  '✓ Fallback机制保障服务可用性'
];

modelCaps.forEach((cap, i) => {
  slide12.addText(cap, {
    x: 0.6, y: 2.2 + i * 0.33, w: 5.4, h: 0.3,
    fontSize: 11, color: COLORS.mediumGray,
    fontFace: 'Microsoft YaHei'
  });
});

// 回答示例
slide12.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 6.4, y: 1.65, w: 3.2, h: 2.5,
  fill: { color: 'FFF7ED' },
  line: { color: COLORS.accent, width: 1 },
  rectRadius: 0.08
});

slide12.addText('💬 回答示例', {
  x: 6.5, y: 1.75, w: 3.0, h: 0.35,
  fontSize: 12, bold: true, color: COLORS.warning,
  fontFace: 'Microsoft YaHei'
});

slide12.addText('学生：国家助学金怎么申请？\n\n🎓 国家助学金申请指南\n\n📋 申请条件：\n家庭经济困难、学习良好、无违纪\n\n💰 资助标准：\n3000-5000元/年\n\n📝 申请流程：\n9月系统申请 → 准备材料 → 辅导员审核 → 学院公示3天 → 10月发放', {
  x: 6.5, y: 2.15, w: 3.0, h: 1.9,
  fontSize: 9, color: COLORS.darkGray,
  fontFace: 'Microsoft YaHei'
});

// 服务稳定性
slide12.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 4.3, w: 9.2, h: 1.2,
  fill: { color: COLORS.lightBlue },
  rectRadius: 0.06
});

slide12.addText('🛡️ 服务保障：多级Fallback机制', {
  x: 0.5, y: 4.4, w: 9.0, h: 0.35,
  fontSize: 12, bold: true, color: COLORS.secondary,
  fontFace: 'Microsoft YaHei'
});

slide12.addText('当Grok-2 API不可用时 → 自动切换到本地RAG直答模式 → 即使API完全故障，知识库内容仍可直接返回，确保服务不中断', {
  x: 0.5, y: 4.8, w: 9.0, h: 0.6,
  fontSize: 11, color: COLORS.mediumGray,
  fontFace: 'Microsoft YaHei'
});

// =============================================
// 第13页：核心功能 - 转人工
// =============================================
let slide13 = pptx.addSlide();
slide13.background = { color: COLORS.lightBg };

slide13.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide13.addText('04', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide13.addText('核心功能详解', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

slide13.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 1.0, w: 9.2, h: 0.5,
  fill: { color: COLORS.secondary },
  rectRadius: 0.05
});

slide13.addText('📞 功能四：转人工服务', {
  x: 0.5, y: 1.0, w: 9.0, h: 0.5,
  fontSize: 16, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', valign: 'middle'
});

// 三大场景
slide13.addText('触发转人工的三大场景', {
  x: 0.4, y: 1.65, w: 9, h: 0.35,
  fontSize: 13, bold: true, color: COLORS.darkGray,
  fontFace: 'Microsoft YaHei'
});

const transferScenarios = [
  { icon: '🗣️', title: '学生主动要求', desc: '学生回复"转人工"、"找老师"等关键词' },
  { icon: '🤔', title: 'AI置信度低', desc: '知识库无匹配结果或匹配度低于阈值' },
  { icon: '⚠️', title: '复杂问题', desc: '涉及多个部门或需要人工判断的综合性问题' }
];

transferScenarios.forEach((ts, i) => {
  const x = 0.4 + i * 3.1;

  slide13.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 2.05, w: 2.95, h: 1.1,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.06),
    rectRadius: 0.06
  });

  slide13.addText(ts.icon, {
    x: x + 0.1, y: 2.15, w: 0.5, h: 0.45,
    fontSize: 20
  });

  slide13.addText(ts.title, {
    x: x + 0.6, y: 2.15, w: 2.2, h: 0.35,
    fontSize: 12, bold: true, color: COLORS.darkGray,
    fontFace: 'Microsoft YaHei'
  });

  slide13.addText(ts.desc, {
    x: x + 0.6, y: 2.5, w: 2.2, h: 0.55,
    fontSize: 10, color: COLORS.mediumGray,
    fontFace: 'Microsoft YaHei'
  });
});

// 老师配置
slide13.addText('已配置老师联系方式（3位）', {
  x: 0.4, y: 3.3, w: 9, h: 0.35,
  fontSize: 13, bold: true, color: COLORS.darkGray,
  fontFace: 'Microsoft YaHei'
});

const teachers = [
  { name: '张老师', dept: '学生工作办公室', specialty: '奖助学金,助学贷款', phone: '010-12345678' },
  { name: '李老师', dept: '宿舍管理办公室', specialty: '宿舍调整,退宿,住宿申请', phone: '010-23456789' },
  { name: '王老师', dept: '教务办公室', specialty: '消三比,成绩单,学籍证明', phone: '010-34567890' }
];

teachers.forEach((t, i) => {
  const x = 0.4 + i * 3.1;

  slide13.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 3.7, w: 2.95, h: 1.7,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.06),
    rectRadius: 0.06
  });

  slide13.addShape(pptx.shapes.OVAL, {
    x: x + 1.1, y: 3.82, w: 0.7, h: 0.7,
    fill: { color: COLORS.secondary }
  });

  slide13.addText(t.name, {
    x: x, y: 3.85, w: 2.95, h: 0.6,
    fontSize: 14, bold: true, color: COLORS.darkGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide13.addText(t.dept, {
    x: x + 0.1, y: 4.6, w: 2.75, h: 0.3,
    fontSize: 10, color: COLORS.mediumGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide13.addText('📱 ' + t.phone, {
    x: x, y: 4.9, w: 2.95, h: 0.25,
    fontSize: 10, color: COLORS.secondary,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide13.addText(t.specialty, {
    x: x + 0.1, y: 5.15, w: 2.75, h: 0.2,
    fontSize: 8, color: COLORS.lightGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
});

// =============================================
// 第14页：核心功能 - 学校公告自动抓取
// =============================================
let slide14 = pptx.addSlide();
slide14.background = { color: COLORS.lightBg };

slide14.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide14.addText('04', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide14.addText('核心功能详解', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

slide14.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 1.0, w: 9.2, h: 0.5,
  fill: { color: COLORS.secondary },
  rectRadius: 0.05
});

slide14.addText('📢 功能五：学校公告自动抓取', {
  x: 0.5, y: 1.0, w: 9.0, h: 0.5,
  fontSize: 16, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', valign: 'middle'
});

// 功能说明
slide14.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 1.65, w: 5.5, h: 2.3,
  fill: { color: COLORS.cardBg },
  shadow: makeShadow(0.06),
  rectRadius: 0.08
});

slide14.addText('核心能力', {
  x: 0.6, y: 1.75, w: 5.1, h: 0.35,
  fontSize: 13, bold: true, color: COLORS.primary,
  fontFace: 'Microsoft YaHei'
});

const fetchCaps = [
  '📰 定时从学校官网、学工处、研究生院等抓取最新公告',
  '🔄 支持多种数据源：RSS订阅、HTTP抓取、JSON接口',
  '📂 自动分类：奖助学金/宿舍/教务/就业等类别',
  '🧹 自动去重：基于标题+时间戳去重，避免重复入库',
  '⏰ 可配置抓取频率：每小时/每天/每周',
  '📊 日志记录：每次抓取结果可查、可追溯'
];

fetchCaps.forEach((cap, i) => {
  slide14.addText(cap, {
    x: 0.6, y: 2.15 + i * 0.3, w: 5.1, h: 0.28,
    fontSize: 10, color: COLORS.mediumGray,
    fontFace: 'Microsoft YaHei'
  });
});

// 数据导入方式
slide14.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 6.1, y: 1.65, w: 3.5, h: 2.3,
  fill: { color: COLORS.lightBlue },
  rectRadius: 0.08
});

slide14.addText('多种数据导入方式', {
  x: 6.2, y: 1.75, w: 3.3, h: 0.35,
  fontSize: 12, bold: true, color: COLORS.secondary,
  fontFace: 'Microsoft YaHei'
});

const importWays = [
  '📄 文件上传：PDF/Word/TXT/JSON',
  '🔗 URL抓取：网页链接直接导入',
  '📡 RSS订阅：自动订阅并更新',
  '📋 批量导入：文件夹一键批量处理',
  '🖥️ API接口：与其他系统对接',
  '📝 手动录入：后台直接添加文本'
];

importWays.forEach((way, i) => {
  slide14.addText(way, {
    x: 6.2, y: 2.15 + i * 0.3, w: 3.3, h: 0.28,
    fontSize: 10, color: COLORS.mediumGray,
    fontFace: 'Microsoft YaHei'
  });
});

// 知识库管理
slide14.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 4.1, w: 9.2, h: 1.4,
  fill: { color: COLORS.cardBg },
  shadow: makeShadow(0.06),
  rectRadius: 0.08
});

slide14.addText('🗂️ 知识库管理后台', {
  x: 0.6, y: 4.2, w: 8.8, h: 0.35,
  fontSize: 13, bold: true, color: COLORS.primary,
  fontFace: 'Microsoft YaHei'
});

slide14.addText('支持增删改查、可视化配置、版本管理、批量操作、导出备份，老师无需编程即可自主维护知识库内容', {
  x: 0.6, y: 4.6, w: 8.8, h: 0.35,
  fontSize: 11, color: COLORS.mediumGray,
  fontFace: 'Microsoft YaHei'
});

const kbFeatures = ['➕ 添加知识', '✏️ 编辑内容', '🗑️ 删除过期', '📤 导出备份', '📊 统计分析'];
kbFeatures.forEach((f, i) => {
  slide14.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.6 + i * 1.78, y: 5.0, w: 1.68, h: 0.4,
    fill: { color: COLORS.lightBlue },
    rectRadius: 0.15
  });
  slide14.addText(f, {
    x: 0.6 + i * 1.78, y: 5.0, w: 1.68, h: 0.4,
    fontSize: 10, color: COLORS.secondary,
    align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
  });
});

// =============================================
// 第15页：知识库内容
// =============================================
let slide15 = pptx.addSlide();
slide15.background = { color: COLORS.lightBg };

slide15.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide15.addText('05', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide15.addText('知识库内容', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

slide15.addText('已集成的学校办事指南与政策文件', {
  x: 1.0, y: 0.5, w: 6, h: 0.3,
  fontSize: 11, color: '93C5FD',
  fontFace: 'Microsoft YaHei', margin: 0
});

// 知识类别
const knowledgeCategories = [
  {
    category: '奖助学金',
    icon: '🎓',
    color: '059669',
    docs: [
      { title: '国家助学金申请指南', source: '学生手册', content: '申请条件、奖励标准、申请流程、时间节点' },
      { title: '国家奖学金评定办法', source: '学生手册', content: '评定条件、奖励标准8000元/年、申请流程' },
      { title: '国家励志奖学金', source: '学生手册', content: '资助对象5000元/年、申请条件、评审流程' },
      { title: '勤工俭学管理办法', source: '学生手册', content: '岗位申请、工资标准、工作要求' }
    ]
  },
  {
    category: '宿舍管理',
    icon: '🏠',
    color: '2563EB',
    docs: [
      { title: '调换宿舍办理流程', source: '宿舍管理规定', content: '申请条件、办理流程、审批时限' },
      { title: '退宿办理指南', source: '宿舍管理规定', content: '毕业退宿流程、押金退还、注意事项' },
      { title: '住宿申请须知', source: '宿舍管理规定', content: '新生申请、老生续租、选房规则' }
    ]
  },
  {
    category: '消三比',
    icon: '✅',
    color: '7C3AED',
    docs: [
      { title: '消三比达标标准', source: '学生手册', content: '学业进度80%+、综素活动10次+、宿舍80分+' },
      { title: '消三比预警消除流程', source: '学生手册', content: '学业预警/综素预警/宿舍预警各自消除流程' }
    ]
  }
];

let kbY = 1.05;
knowledgeCategories.forEach((kc, catIdx) => {
  // 类别标题
  slide15.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: kbY, w: 9.2, h: 0.4,
    fill: { color: kc.color },
    rectRadius: 0.05
  });

  slide15.addText(kc.icon + ' ' + kc.category, {
    x: 0.5, y: kbY, w: 9.0, h: 0.4,
    fontSize: 13, bold: true, color: COLORS.white,
    fontFace: 'Microsoft YaHei', valign: 'middle'
  });

  kbY += 0.45;

  // 文档列表
  kc.docs.forEach((doc, docIdx) => {
    const docX = 0.4 + (docIdx % 2) * 4.6;
    const docY = kbY + Math.floor(docIdx / 2) * 0.85;

    slide15.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: docX, y: docY, w: 4.4, h: 0.75,
      fill: { color: COLORS.cardBg },
      shadow: makeShadow(0.04),
      rectRadius: 0.05
    });

    slide15.addText(doc.title, {
      x: docX + 0.15, y: docY + 0.08, w: 4.1, h: 0.3,
      fontSize: 11, bold: true, color: COLORS.darkGray,
      fontFace: 'Microsoft YaHei'
    });

    slide15.addText('📌 ' + doc.content, {
      x: docX + 0.15, y: docY + 0.38, w: 4.1, h: 0.3,
      fontSize: 9, color: COLORS.mediumGray,
      fontFace: 'Microsoft YaHei'
    });
  });

  kbY += Math.ceil(kc.docs.length / 2) * 0.9 + 0.15;
});

// =============================================
// 第16页：知识库内容（续）
// =============================================
let slide16 = pptx.addSlide();
slide16.background = { color: COLORS.lightBg };

slide16.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide16.addText('05', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide16.addText('知识库内容（续）', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

slide16.addText('其他办事指南与政策文件', {
  x: 1.0, y: 0.5, w: 6, h: 0.3,
  fontSize: 11, color: '93C5FD',
  fontFace: 'Microsoft YaHei', margin: 0
});

// 续期知识库
const knowledgeCategories2 = [
  {
    category: '其他校园事务',
    icon: '📄',
    color: 'DC2626',
    docs: [
      { title: '请假制度', source: '学生手册', content: '请假类型、审批流程、证明材料、补办规定' },
      { title: '成绩单办理', source: '教务处', content: '中英文成绩单、在读证明、学籍证明' },
      { title: '户籍办理', source: '保卫处', content: '户籍迁移、借出、归还流程' },
      { title: '毕业手续', source: '教务处', content: '离校手续、毕业证领取、档案转递' }
    ]
  },
  {
    category: '教务相关',
    icon: '📚',
    color: 'D97706',
    docs: [
      { title: '选课指南', source: '教务处', content: '选课时间、选课规则、补退选流程' },
      { title: '转专业办法', source: '教务处', content: '申请条件、考核流程、时间安排' },
      { title: '休学复学', source: '教务处', content: '休学申请、复学手续、学籍保留' }
    ]
  }
];

let kb2Y = 1.05;
knowledgeCategories2.forEach((kc) => {
  slide16.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: kb2Y, w: 9.2, h: 0.4,
    fill: { color: kc.color },
    rectRadius: 0.05
  });

  slide16.addText(kc.icon + ' ' + kc.category, {
    x: 0.5, y: kb2Y, w: 9.0, h: 0.4,
    fontSize: 13, bold: true, color: COLORS.white,
    fontFace: 'Microsoft YaHei', valign: 'middle'
  });

  kb2Y += 0.45;

  kc.docs.forEach((doc, docIdx) => {
    const docX = 0.4 + (docIdx % 2) * 4.6;
    const docY = kb2Y + Math.floor(docIdx / 2) * 0.85;

    slide16.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: docX, y: docY, w: 4.4, h: 0.75,
      fill: { color: COLORS.cardBg },
      shadow: makeShadow(0.04),
      rectRadius: 0.05
    });

    slide16.addText(doc.title, {
      x: docX + 0.15, y: docY + 0.08, w: 4.1, h: 0.3,
      fontSize: 11, bold: true, color: COLORS.darkGray,
      fontFace: 'Microsoft YaHei'
    });

    slide16.addText('📌 ' + doc.content, {
      x: docX + 0.15, y: docY + 0.38, w: 4.1, h: 0.3,
      fontSize: 9, color: COLORS.mediumGray,
      fontFace: 'Microsoft YaHei'
    });
  });

  kb2Y += Math.ceil(kc.docs.length / 2) * 0.9 + 0.15;
});

// 知识库统计
slide16.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 4.3, w: 9.2, h: 1.15,
  fill: { color: COLORS.primary },
  rectRadius: 0.08
});

const kbStats = [
  { val: '5', label: '知识类别' },
  { val: '20+', label: '文档数量' },
  { val: '3', label: '老师配置' },
  { val: '7×24', label: '服务时间' }
];

kbStats.forEach((s, i) => {
  const sx = 0.7 + i * 2.3;
  slide16.addText(s.val, {
    x: sx, y: 4.4, w: 2.0, h: 0.55,
    fontSize: 28, bold: true, color: COLORS.accent,
    align: 'center', fontFace: 'Arial'
  });
  slide16.addText(s.label, {
    x: sx, y: 4.95, w: 2.0, h: 0.4,
    fontSize: 12, color: COLORS.white,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
});

// =============================================
// 第17页：运营保障
// =============================================
let slide17 = pptx.addSlide();
slide17.background = { color: COLORS.lightBg };

slide17.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide17.addText('06', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide17.addText('运营保障', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 三大保障
const guarantees = [
  {
    icon: '🔄',
    title: '转人工兜底机制',
    color: 'DC2626',
    points: [
      'AI无法解答时自动转人工，不推诿',
      '3位老师按专长分工，覆盖主要事务',
      '转人工消息包含姓名+电话+部门+专长',
      '工作时间电话即拨即通',
      '非工作时间引导至值班老师或工单'
    ]
  },
  {
    icon: '🛡️',
    title: '服务稳定性保障',
    color: '059669',
    points: [
      'FastAPI异步架构，支持高并发',
      '多级Fallback：API→RAG直答→知识库',
      'ChromaDB本地持久化，数据不丢失',
      '系统日志完整，可追溯每条问答',
      '健康检查接口，实时监控服务状态'
    ]
  },
  {
    icon: '🔒',
    title: '内容安全与合规',
    color: '2563EB',
    points: [
      '回答内容基于学校官方文件，准确可信',
      '敏感词过滤机制，防止不当内容',
      '知识库权限分级，管理员/普通用户',
      '用户数据隔离，不泄露学生隐私',
      '微信官方API接入，安全合规'
    ]
  }
];

guarantees.forEach((g, i) => {
  const x = 0.4 + i * 3.1;

  slide17.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.1, w: 2.95, h: 3.9,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.08),
    rectRadius: 0.08
  });

  // 图标圆
  slide17.addShape(pptx.shapes.OVAL, {
    x: x + 1.07, y: 1.25, w: 0.8, h: 0.8,
    fill: { color: g.color }
  });

  slide17.addText(g.icon, {
    x: x + 1.07, y: 1.3, w: 0.8, h: 0.7,
    fontSize: 28, align: 'center', valign: 'middle'
  });

  slide17.addText(g.title, {
    x: x + 0.1, y: 2.15, w: 2.75, h: 0.45,
    fontSize: 13, bold: true, color: COLORS.darkGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  g.points.forEach((p, j) => {
    slide17.addText('• ' + p, {
      x: x + 0.15, y: 2.65 + j * 0.42, w: 2.65, h: 0.4,
      fontSize: 10, color: COLORS.mediumGray,
      fontFace: 'Microsoft YaHei'
    });
  });
});

// 服务可用性承诺
slide17.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 5.1, w: 9.2, h: 0.4,
  fill: { color: 'F0FDF4' },
  line: { color: COLORS.success, width: 1 },
  rectRadius: 0.05
});

slide17.addText('✅ 服务可用性承诺：系统故障2小时内响应，知识库更新2工作日内完成，技术支持7×24小时在线', {
  x: 0.5, y: 5.1, w: 9.0, h: 0.4,
  fontSize: 11, color: COLORS.success,
  fontFace: 'Microsoft YaHei', valign: 'middle'
});

// =============================================
// 第18页：效果预估
// =============================================
let slide18 = pptx.addSlide();
slide18.background = { color: COLORS.lightBg };

slide18.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide18.addText('07', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide18.addText('效果预估', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

slide18.addText('效率提升与成本节省测算', {
  x: 1.0, y: 0.5, w: 6, h: 0.3,
  fontSize: 11, color: '93C5FD',
  fontFace: 'Microsoft YaHei', margin: 0
});

// 效率提升数据
const efficiencyData = [
  { metric: '80%+', label: '咨询自动化率', desc: '常见问题由AI直接解答，无需人工介入' },
  { metric: '95%+', label: '学生满意度', desc: '即时响应+精准答案，学生体验显著提升' },
  { metric: '60%', label: '人力节省', desc: '辅导员日常咨询工作量减少六成，可专注复杂事务' },
  { metric: '3秒', label: '平均响应时间', desc: '学生提问后3秒内获得答案，告别等待' }
];

efficiencyData.forEach((e, i) => {
  const x = 0.4 + (i % 2) * 4.7;
  const y = 1.1 + Math.floor(i / 2) * 1.55;

  slide18.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: 4.5, h: 1.4,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.08),
    rectRadius: 0.08
  });

  slide18.addText(e.metric, {
    x: x + 0.15, y: y + 0.15, w: 1.6, h: 0.8,
    fontSize: 32, bold: true, color: COLORS.secondary,
    fontFace: 'Arial'
  });

  slide18.addText(e.label, {
    x: x + 1.8, y: y + 0.2, w: 2.5, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.darkGray,
    fontFace: 'Microsoft YaHei'
  });

  slide18.addText(e.desc, {
    x: x + 1.8, y: y + 0.65, w: 2.5, h: 0.6,
    fontSize: 10, color: COLORS.mediumGray,
    fontFace: 'Microsoft YaHei'
  });
});

// 成本节省测算
slide18.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 4.35, w: 9.2, h: 1.15,
  fill: { color: COLORS.primary },
  rectRadius: 0.08
});

slide18.addText('💰 成本节省测算（以10000名学生规模为例）', {
  x: 0.5, y: 4.45, w: 9.0, h: 0.4,
  fontSize: 14, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei'
});

const costData = [
  { label: '传统模式年成本', val: '¥50万+' },
  { label: 'AI辅导员年成本', val: '¥9.4万' },
  { label: '年度节省', val: '¥40万+' },
  { label: '投资回报周期', val: '3-6个月' }
];

costData.forEach((c, i) => {
  const cx = 0.6 + i * 2.3;
  slide18.addText(c.val, {
    x: cx, y: 4.85, w: 2.1, h: 0.35,
    fontSize: 16, bold: true, color: COLORS.accent,
    align: 'center', fontFace: 'Arial'
  });
  slide18.addText(c.label, {
    x: cx, y: 5.2, w: 2.1, h: 0.25,
    fontSize: 10, color: '93C5FD',
    align: 'center', fontFace: 'Microsoft YaHei'
  });
});

// =============================================
// 第19页：预算方案
// =============================================
let slide19 = pptx.addSlide();
slide19.background = { color: COLORS.lightBg };

slide19.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide19.addText('08', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide19.addText('预算方案', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

slide19.addText('三档方案，满足不同规模与预算需求', {
  x: 1.0, y: 0.5, w: 6, h: 0.3,
  fontSize: 11, color: '93C5FD',
  fontFace: 'Microsoft YaHei', margin: 0
});

// 三档方案（基于实际API价格和token使用量测算）
// 测算基础：6,000月活用户 × 15次查询/月 × 1,500 tokens/查询
// 参考价格：通义千问Plus ¥3.65/输入M + ¥14.6/输出M，Gemini Flash ¥0.55/输入M + ¥2.19/输出M
const budgets = [
  {
    name: '保守型',
    price: '¥2,555',
    period: '季度',
    color: '059669',
    lightColor: 'D1FAE5',
    target: '小型院校/试点',
    features: [
      'Gemini 1.5 Flash API',
      '基础功能部署',
      '标准知识库配置',
      '1个微信公众号接入',
      '基础技术支持',
      '月度知识库更新',
      'Q&A报表'
    ],
    limits: [
      '不支持私有化部署',
      '不支持定制开发',
      '不提供现场服务'
    ]
  },
  {
    name: '均衡型',
    price: '¥23,530',
    period: '季度',
    color: '2563EB',
    lightColor: 'DBEAFE',
    target: '中型院校/标准需求',
    popular: true,
    features: [
      '通义千问Plus API（中文最优）',
      '完整功能部署',
      '高级知识库配置',
      '3个公众号接入',
      '优先技术支持',
      '周度知识库更新',
      '详细数据报表',
      '转人工工单系统',
      '学校公告自动抓取'
    ],
    limits: [
      '不支持私有化部署',
      '有限定制开发'
    ]
  },
  {
    name: '激进型',
    price: '¥29,783',
    period: '季度',
    color: '7C3AED',
    lightColor: 'EDE9FE',
    target: '大型院校/深度需求',
    features: [
      'Kimi API（国产高性价比）',
      '全功能+私有化部署',
      '无限公众号接入',
      '7×24专属技术支持',
      '实时知识库更新',
      '高级数据分析平台',
      '多角色权限管理',
      '与学校系统对接',
      'AI模型微调优化',
      '季度现场培训'
    ],
    limits: []
  }
];

budgets.forEach((b, i) => {
  const x = 0.3 + i * 3.23;
  const isPopular = b.popular;

  slide19.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.0, w: 3.1, h: 4.5,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(isPopular ? 0.15 : 0.08),
    rectRadius: 0.1
  });

  if (isPopular) {
    slide19.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.8, y: 0.8, w: 1.5, h: 0.35,
      fill: { color: COLORS.accent },
      rectRadius: 0.15
    });
    slide19.addText('推荐', {
      x: x + 0.8, y: 0.8, w: 1.5, h: 0.35,
      fontSize: 11, bold: true, color: COLORS.white,
      align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
    });
  }

  slide19.addText(b.name, {
    x: x, y: 1.15, w: 3.1, h: 0.4,
    fontSize: 16, bold: true, color: b.color,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide19.addText(b.target, {
    x: x, y: 1.55, w: 3.1, h: 0.3,
    fontSize: 10, color: COLORS.lightGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide19.addText(b.price, {
    x: x, y: 1.85, w: 3.1, h: 0.6,
    fontSize: 26, bold: true, color: b.color,
    align: 'center', fontFace: 'Arial'
  });

  slide19.addText(b.period, {
    x: x, y: 2.45, w: 3.1, h: 0.25,
    fontSize: 10, color: COLORS.lightGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide19.addShape(pptx.shapes.RECTANGLE, {
    x: x + 0.3, y: 2.75, w: 2.5, h: 0.02,
    fill: { color: COLORS.border }
  });

  b.features.forEach((f, j) => {
    slide19.addText('✓ ' + f, {
      x: x + 0.15, y: 2.85 + j * 0.32, w: 2.8, h: 0.3,
      fontSize: 10, color: COLORS.mediumGray,
      fontFace: 'Microsoft YaHei'
    });
  });

  if (b.limits.length > 0) {
    b.limits.forEach((l, j) => {
      slide19.addText('✗ ' + l, {
        x: x + 0.15, y: 2.85 + (b.features.length + j) * 0.32, w: 2.8, h: 0.3,
        fontSize: 9, color: COLORS.lightGray,
        fontFace: 'Microsoft YaHei'
      });
    });
  }
});

// 价格说明
slide19.addText('※ 价格包含成本+服务费（服务费比例30%）', {
  x: 0.4, y: 5.15, w: 9.2, h: 0.35,
  fontSize: 11, color: COLORS.mediumGray,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// =============================================
// 第19b页：费用构成明细
// =============================================
let slide19b = pptx.addSlide();
slide19b.background = { color: COLORS.lightBg };

slide19b.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide19b.addText('08', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide19b.addText('费用构成明细', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

slide19b.addText('本方案基于6,000月活用户测算（15次查询/月 × 1,500 tokens/查询）', {
  x: 1.0, y: 0.5, w: 8, h: 0.3,
  fontSize: 11, color: '93C5FD',
  fontFace: 'Microsoft YaHei', margin: 0
});

// 费用明细表格 - 四列布局（总宽度10.0，适配幻灯片宽度10英寸）
const col1X = 0.3;      // 行标签列
const col2X = 2.2;      // 保守型列
const col3X = 4.9;      // 均衡型列
const col4X = 7.6;      // 激进型列
const col1W = 1.8;      // 行标签宽度
const colW = 2.35;      // 三档方案列宽度 (7.6+2.35=9.95<10, OK)
const rowH = 0.42;
const startY = 1.05;

// 表头
const headers = [
  { x: col1X, w: col1W, label: '费用项', color: COLORS.primary, textColor: COLORS.white },
  { x: col2X, w: colW, label: '保守型', color: '059669', textColor: COLORS.white },
  { x: col3X, w: colW, label: '均衡型', color: '2563EB', textColor: COLORS.white },
  { x: col4X, w: colW, label: '激进型', color: '7C3AED', textColor: COLORS.white }
];

headers.forEach(h => {
  slide19b.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: h.x, y: startY, w: h.w, h: rowH,
    fill: { color: h.color },
    rectRadius: 0.05
  });
  slide19b.addText(h.label, {
    x: h.x, y: startY, w: h.w, h: rowH,
    fontSize: 12, bold: true, color: h.textColor,
    align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
  });
});

// 表格数据行
const costRows = [
  { label: 'API调用费', values: ['¥405', '¥11,640', '¥16,200'] },
  { label: '云服务器+带宽', values: ['¥1,350', '¥1,950', '¥1,950'] },
  { label: '运营维护', values: ['¥210', '¥510', '¥760'] },
  { label: '知识库更新费', values: ['¥0', '¥2,000', '¥2,000'] },
  { label: '技术支持费', values: ['¥0', '¥2,000', '¥2,000'] },
  { label: '成本合计', values: ['¥1,965', '¥18,100', '¥22,910'], isTotal: true },
  { label: '服务费（30%）', values: ['¥590', '¥5,430', '¥6,873'], isService: true },
  { label: '对外报价', values: ['¥2,555', '¥23,530', '¥29,783'], isPrice: true }
];

costRows.forEach((row, ri) => {
  const y = startY + rowH + ri * rowH;
  const isAlt = ri % 2 === 1;
  
  // 左侧行标签
  let labelBg = isAlt ? COLORS.lightBg : COLORS.cardBg;
  slide19b.addShape(pptx.shapes.RECTANGLE, {
    x: col1X, y: y, w: col1W, h: rowH,
    fill: { color: labelBg },
    line: { color: COLORS.border, width: 0.5 }
  });
  slide19b.addText(row.label, {
    x: col1X, y: y, w: col1W, h: rowH,
    fontSize: 9, color: COLORS.darkGray,
    align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
  });
  
  // 三个方案的数值
  const colX = [col2X, col3X, col4X];
  const colors = ['059669', '2563EB', '7C3AED'];
  
  row.values.forEach((val, ci) => {
    let bgColor = isAlt ? COLORS.lightBg : COLORS.cardBg;
    let textColor = COLORS.darkGray;
    let bold = false;
    
    if (row.isTotal) { bgColor = COLORS.lightBlue; bold = true; }
    if (row.isService) { bgColor = 'FEF3C7'; textColor = COLORS.warning; bold = true; }
    if (row.isPrice) { bgColor = colors[ci]; textColor = COLORS.white; bold = true; }
    
    slide19b.addShape(pptx.shapes.RECTANGLE, {
      x: colX[ci], y: y, w: colW, h: rowH,
      fill: { color: bgColor },
      line: { color: COLORS.border, width: 0.5 }
    });
    slide19b.addText(val, {
      x: colX[ci], y: y, w: colW, h: rowH,
      fontSize: row.isPrice ? 12 : 10, bold: bold, color: textColor,
      align: 'center', valign: 'middle', fontFace: 'Arial'
    });
  });
});

// 图例说明
slide19b.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.3, y: 5.0, w: 9.4, h: 0.5,
  fill: { color: COLORS.cardBg },
  rectRadius: 0.05
});

slide19b.addText('※ 服务费为我们赚取的利润（30%），用于持续优化服务与技术研发 | 成本明细公开透明', {
  x: 0.4, y: 5.0, w: 9.2, h: 0.5,
  fontSize: 10, color: COLORS.mediumGray,
  align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
});

// =============================================
// 第20页：实施计划
// =============================================
let slide20 = pptx.addSlide();
slide20.background = { color: COLORS.lightBg };

slide20.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide20.addText('09', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide20.addText('实施计划', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

slide20.addText('三阶段里程碑，整体周期约6-8周', {
  x: 1.0, y: 0.5, w: 6, h: 0.3,
  fontSize: 11, color: '93C5FD',
  fontFace: 'Microsoft YaHei', margin: 0
});

// 三个阶段
const phases = [
  {
    phase: '第一阶段',
    title: '部署与适配',
    duration: '第1-2周',
    color: '2563EB',
    tasks: [
      { task: '服务器环境准备', desc: '云服务器或本地服务器部署准备' },
      { task: '微信公众号配置', desc: 'AppID/Secret/Token/AESKey配置' },
      { task: '知识库初始化', desc: '学校文件PDF/Word解析入库' },
      { task: '基础功能测试', desc: '问答、转人工、意图识别测试' }
    ]
  },
  {
    phase: '第二阶段',
    title: '定制与优化',
    duration: '第3-5周',
    color: '7C3AED',
    tasks: [
      { task: '知识库内容扩充', desc: '根据学校情况补充更多办事指南' },
      { task: '老师联系方式配置', desc: '按实际部门配置转接规则' },
      { task: '回答质量优化', desc: '调优Prompt，提升回答准确率' },
      { task: '内测与反馈', desc: '邀请辅导员和学生内测，收集反馈' }
    ]
  },
  {
    phase: '第三阶段',
    title: '上线与运营',
    duration: '第6-8周',
    color: '059669',
    tasks: [
      { task: '正式上线发布', desc: '全量开放给全校学生使用' },
      { task: '持续知识库运营', desc: '定期更新文件、补充新内容' },
      { task: '效果跟踪分析', desc: '统计问答量、满意度、转人工率' },
      { task: '迭代优化', desc: '基于数据持续优化系统效果' }
    ]
  }
];

phases.forEach((p, i) => {
  const x = 0.35 + i * 3.2;

  slide20.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.1, w: 3.05, h: 4.4,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.08),
    rectRadius: 0.1
  });

  // 阶段标签
  slide20.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.8, y: 1.0, w: 1.45, h: 0.35,
    fill: { color: p.color },
    rectRadius: 0.15
  });

  slide20.addText(p.phase, {
    x: x + 0.8, y: 1.0, w: 1.45, h: 0.35,
    fontSize: 11, bold: true, color: COLORS.white,
    align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
  });

  slide20.addText(p.title, {
    x: x, y: 1.4, w: 3.05, h: 0.4,
    fontSize: 15, bold: true, color: COLORS.darkGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide20.addText(p.duration, {
    x: x, y: 1.8, w: 3.05, h: 0.3,
    fontSize: 11, color: p.color,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  p.tasks.forEach((t, j) => {
    slide20.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.1, y: 2.2 + j * 0.85, w: 2.85, h: 0.75,
      fill: { color: COLORS.lightBg },
      rectRadius: 0.05
    });

    slide20.addText('▸ ' + t.task, {
      x: x + 0.18, y: 2.25 + j * 0.85, w: 2.7, h: 0.35,
      fontSize: 11, bold: true, color: COLORS.darkGray,
      fontFace: 'Microsoft YaHei'
    });

    slide20.addText(t.desc, {
      x: x + 0.18, y: 2.55 + j * 0.85, w: 2.7, h: 0.35,
      fontSize: 9, color: COLORS.lightGray,
      fontFace: 'Microsoft YaHei'
    });
  });
});

// =============================================
// 第21页：合作方式
// =============================================
let slide21 = pptx.addSlide();
slide21.background = { color: COLORS.lightBg };

slide21.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.85,
  fill: { color: COLORS.primary }
});

slide21.addText('10', {
  x: 0.4, y: 0.15, w: 0.6, h: 0.55,
  fontSize: 28, bold: true, color: COLORS.accent,
  fontFace: 'Arial'
});

slide21.addText('合作方式', {
  x: 1.0, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 落地模式
slide21.addText('落地模式', {
  x: 0.4, y: 1.1, w: 9, h: 0.4,
  fontSize: 14, bold: true, color: COLORS.darkGray,
  fontFace: 'Microsoft YaHei'
});

const deployModes = [
  { icon: '☁️', title: 'SaaS云服务', desc: '我方托管在云端，学校无需服务器，快速接入', suitable: '适合：快速上线、无IT团队的院校' },
  { icon: '🖥️', title: '私有化部署', desc: '部署在学校自有服务器，数据完全自主掌控', suitable: '适合：对数据安全有要求、有IT运维团队的院校' },
  { icon: '🔀', title: '混合部署', desc: '核心数据本地，知识库云端，兼顾安全与灵活', suitable: '适合：大型院校、特殊合规要求的院校' }
];

deployModes.forEach((dm, i) => {
  const x = 0.4 + i * 3.1;

  slide21.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.55, w: 2.95, h: 2.2,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.06),
    rectRadius: 0.08
  });

  slide21.addText(dm.icon, {
    x: x, y: 1.65, w: 2.95, h: 0.6,
    fontSize: 28, align: 'center'
  });

  slide21.addText(dm.title, {
    x: x, y: 2.25, w: 2.95, h: 0.4,
    fontSize: 13, bold: true, color: COLORS.darkGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide21.addText(dm.desc, {
    x: x + 0.1, y: 2.65, w: 2.75, h: 0.55,
    fontSize: 10, color: COLORS.mediumGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide21.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.15, y: 3.25, w: 2.65, h: 0.4,
    fill: { color: COLORS.lightBlue },
    rectRadius: 0.15
  });

  slide21.addText(dm.suitable, {
    x: x + 0.15, y: 3.25, w: 2.65, h: 0.4,
    fontSize: 8, color: COLORS.secondary,
    align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
  });
});

// 服务保障
slide21.addText('服务保障', {
  x: 0.4, y: 3.95, w: 9, h: 0.4,
  fontSize: 14, bold: true, color: COLORS.darkGray,
  fontFace: 'Microsoft YaHei'
});

const serviceGuards = [
  { icon: '📅', title: '服务周期', desc: '合同期1年起，期满自动续约' },
  { icon: '🔧', title: '运维支持', desc: '7×24小时技术支持，故障2小时内响应' },
  { icon: '📚', title: '知识库运营', desc: '协助知识库更新维护，不额外收费' },
  { icon: '🎓', title: '培训服务', desc: '提供操作培训视频和文档' }
];

serviceGuards.forEach((sg, i) => {
  const x = 0.4 + i * 2.35;

  slide21.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 4.4, w: 2.2, h: 1.1,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.05),
    rectRadius: 0.06
  });

  slide21.addText(sg.icon, {
    x: x, y: 4.5, w: 2.2, h: 0.4,
    fontSize: 18, align: 'center'
  });

  slide21.addText(sg.title, {
    x: x, y: 4.9, w: 2.2, h: 0.25,
    fontSize: 10, bold: true, color: COLORS.darkGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide21.addText(sg.desc, {
    x: x + 0.1, y: 5.15, w: 2.0, h: 0.3,
    fontSize: 8, color: COLORS.mediumGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
});

// =============================================
// 第22页：总结与展望
// =============================================
let slide22 = pptx.addSlide();
slide22.background = { color: COLORS.primary };

// 装饰
slide22.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.06,
  fill: { color: COLORS.accent }
});

slide22.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0.06, w: 0.12, h: 3,
  fill: { color: COLORS.secondary, transparency: 40 }
});

slide22.addText('总结与展望', {
  x: 0.5, y: 0.4, w: 9, h: 0.7,
  fontSize: 32, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei'
});

// 核心价值回顾
const summaryPoints = [
  '✅ 7×24即时响应，彻底解决学生"等待难"问题',
  '✅ RAG知识库精准解答，答案基于学校官方文件',
  '✅ 5大意图全覆盖，奖助学金/宿舍/消三比等全涵盖',
  '✅ 转人工兜底机制，服务闭环有保障',
  '✅ 微信公众号入口，学生无需下载，即开即用'
];

summaryPoints.forEach((sp, i) => {
  slide22.addText(sp, {
    x: 0.6, y: 1.3 + i * 0.55, w: 8.8, h: 0.5,
    fontSize: 15, color: COLORS.white,
    fontFace: 'Microsoft YaHei'
  });
});

// 未来展望
slide22.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 4.15, w: 9.2, h: 1.3,
  fill: { color: COLORS.white, transparency: 90 },
  rectRadius: 0.1
});

slide22.addText('🚀 未来升级方向', {
  x: 0.5, y: 4.25, w: 9.0, h: 0.4,
  fontSize: 14, bold: true, color: COLORS.accent,
  fontFace: 'Microsoft YaHei'
});

const futurePoints = ['多模态交互（图片/语音）', 'Agent智能体自动化办事', '与教务/宿管系统深度对接', '学生画像与个性化服务'];
futurePoints.forEach((fp, i) => {
  slide22.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.5 + i * 2.25, y: 4.7, w: 2.15, h: 0.6,
    fill: { color: COLORS.accent, transparency: 80 },
    rectRadius: 0.1
  });
  slide22.addText(fp, {
    x: 0.5 + i * 2.25, y: 4.7, w: 2.15, h: 0.6,
    fontSize: 11, color: COLORS.white,
    align: 'center', valign: 'middle', fontFace: 'Microsoft YaHei'
  });
});

// =============================================
// 第23页：联系我们
// =============================================
let slide23 = pptx.addSlide();
slide23.background = { color: COLORS.lightBg };

slide23.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 2.8,
  fill: { color: COLORS.primary }
});

slide23.addText('联系我们', {
  x: 0.5, y: 0.5, w: 9, h: 0.8,
  fontSize: 36, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei'
});

slide23.addText('期待与贵校携手，共建智慧校园服务新体验', {
  x: 0.5, y: 1.3, w: 9, h: 0.5,
  fontSize: 16, color: '93C5FD',
  fontFace: 'Microsoft YaHei'
});

slide23.addText('AI辅导员项目组', {
  x: 0.5, y: 1.9, w: 9, h: 0.4,
  fontSize: 14, color: 'BFDBFE',
  fontFace: 'Microsoft YaHei'
});

// 联系信息卡
const contactInfo = [
  { icon: '📧', label: '邮箱', val: '3977871701@qq.com' },
  { icon: '📞', label: '电话', val: '13825850878' },
  { icon: '🌐', label: '官网', val: '待定' },
  { icon: '💬', label: '微信', val: '13825850878' }
];

contactInfo.forEach((c, i) => {
  const x = 0.4 + i * 2.35;

  slide23.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 3.1, w: 2.2, h: 1.5,
    fill: { color: COLORS.cardBg },
    shadow: makeShadow(0.06),
    rectRadius: 0.08
  });

  slide23.addText(c.icon, {
    x: x, y: 3.2, w: 2.2, h: 0.55,
    fontSize: 26, align: 'center'
  });

  slide23.addText(c.label, {
    x: x, y: 3.75, w: 2.2, h: 0.3,
    fontSize: 11, color: COLORS.lightGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide23.addText(c.val, {
    x: x, y: 4.05, w: 2.2, h: 0.4,
    fontSize: 11, color: COLORS.darkGray,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
});

// 底部信息
slide23.addText('© 2026 AI辅导员项目组 · 保留所有权利', {
  x: 0.5, y: 5.15, w: 9, h: 0.3,
  fontSize: 10, color: COLORS.lightGray,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// =============================================
// 保存文件
// =============================================
pptx.writeFile({ fileName: '/Users/xylei/work/projects/ai-tutor-ppt/学院AI辅导员客户版PPT.pptx' })
  .then(() => {
    console.log('✅ 客户版PPT生成成功！');
    console.log('📍 文件路径: /Users/xylei/work/projects/ai-tutor-ppt/学院AI辅导员客户版PPT.pptx');
    console.log('📊 共计 23 页幻灯片');
  })
  .catch(err => {
    console.error('❌ PPT生成失败:', err);
  });
