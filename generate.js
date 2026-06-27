const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();

// 设置演示文稿属性
pptx.author = 'AI辅导员团队';
pptx.title = '学院AI辅导员系统介绍';
pptx.subject = '智慧校园服务新体验';

// 配色方案
const COLORS = {
  primary: '2B579A',      // 学院蓝
  secondary: '1E88E5',     // 科技蓝
  accent: 'FF9800',        // 橙色强调
  white: 'FFFFFF',
  lightGray: 'F5F5F5',
  darkGray: '424242',
  textDark: '212121',
  textLight: '757575'
};

// 创建阴影的工厂函数
const makeShadow = () => ({
  type: 'outer',
  color: '000000',
  blur: 4,
  offset: 2,
  angle: 135,
  opacity: 0.15
});

// ========== 第1页：封面 ==========
let slide1 = pptx.addSlide();
slide1.background = { color: COLORS.primary };

// 顶部装饰线
slide1.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0.8, w: 10, h: 0.05,
  fill: { color: COLORS.accent }
});

// 主标题
slide1.addText('学院AI辅导员系统', {
  x: 0.5, y: 1.5, w: 9, h: 1,
  fontSize: 44, bold: true, color: COLORS.white,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// 副标题
slide1.addText('智慧校园服务新体验', {
  x: 0.5, y: 2.5, w: 9, h: 0.8,
  fontSize: 28, color: COLORS.accent,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// 底部装饰线
slide1.addShape(pptx.shapes.RECTANGLE, {
  x: 3, y: 3.5, w: 4, h: 0.03,
  fill: { color: COLORS.white }
});

// 关键词
slide1.addText('🤖 RAG知识库  |  💬 智能问答  |  📢 微信公众号', {
  x: 0.5, y: 4, w: 9, h: 0.5,
  fontSize: 16, color: 'B3D4FC',
  align: 'center', fontFace: 'Microsoft YaHei'
});

// ========== 第2页：项目背景与目标 ==========
let slide2 = pptx.addSlide();
slide2.background = { color: COLORS.lightGray };

// 顶部色块
slide2.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide2.addText('📋 项目背景与目标', {
  x: 0.5, y: 0.15, w: 9, h: 0.6,
  fontSize: 28, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 痛点卡片
slide2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 1.2, w: 4.4, h: 2.0,
  fill: { color: COLORS.white },
  shadow: makeShadow(), rectRadius: 0.08
});

slide2.addText('🎯 痛点', {
  x: 0.6, y: 1.3, w: 4, h: 0.4,
  fontSize: 18, bold: true, color: COLORS.primary,
  fontFace: 'Microsoft YaHei'
});

slide2.addText([
  { text: '• 学生咨询需求分散', options: { breakLine: true } },
  { text: '• 辅导员工作繁重', options: { breakLine: true } },
  { text: '• 回复不及时、标准化程度低', options: {} }
], {
  x: 0.6, y: 1.75, w: 4, h: 1.3,
  fontSize: 14, color: COLORS.textDark,
  fontFace: 'Microsoft YaHei', valign: 'top'
});

// 趋势卡片
slide2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 5.2, y: 1.2, w: 4.4, h: 2.0,
  fill: { color: COLORS.white },
  shadow: makeShadow(), rectRadius: 0.08
});

slide2.addText('📱 趋势', {
  x: 5.4, y: 1.3, w: 4, h: 0.4,
  fontSize: 18, bold: true, color: COLORS.secondary,
  fontFace: 'Microsoft YaHei'
});

slide2.addText([
  { text: '• 高校数字化转型加速', options: { breakLine: true } },
  { text: '• 学生习惯通过微信获取服务', options: { breakLine: true } },
  { text: '• AI技术成熟度已达到实用水平', options: {} }
], {
  x: 5.4, y: 1.75, w: 4, h: 1.3,
  fontSize: 14, color: COLORS.textDark,
  fontFace: 'Microsoft YaHei', valign: 'top'
});

// 解决方案卡片
slide2.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.4, y: 3.4, w: 9.2, h: 1.8,
  fill: { color: COLORS.primary },
  shadow: makeShadow(), rectRadius: 0.08
});

slide2.addText('🤖 解决方案', {
  x: 0.6, y: 3.55, w: 8.8, h: 0.4,
  fontSize: 18, bold: true, color: COLORS.accent,
  fontFace: 'Microsoft YaHei'
});

slide2.addText('基于RAG知识库的智能问答机器人，7×24小时即时响应', {
  x: 0.6, y: 4.0, w: 8.8, h: 0.4,
  fontSize: 16, color: COLORS.white,
  fontFace: 'Microsoft YaHei'
});

slide2.addText('目标：减轻辅导员工作负担，提升学生满意度，实现服务标准化', {
  x: 0.6, y: 4.5, w: 8.8, h: 0.5,
  fontSize: 14, color: 'B3D4FC',
  fontFace: 'Microsoft YaHei'
});

// ========== 第3页：系统架构 ==========
let slide3 = pptx.addSlide();
slide3.background = { color: COLORS.lightGray };

slide3.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide3.addText('🏗️ 系统架构', {
  x: 0.5, y: 0.15, w: 9, h: 0.6,
  fontSize: 28, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 三位一体架构图
const boxY = 1.5;
const boxH = 1.8;
const boxW = 2.8;

// RAG知识库
slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: boxY, w: boxW, h: boxH,
  fill: { color: 'E3F2FD' },
  line: { color: COLORS.secondary, width: 2 },
  rectRadius: 0.08
});

slide3.addText('🧠 RAG知识库', {
  x: 0.5, y: boxY + 0.15, w: boxW, h: 0.4,
  fontSize: 16, bold: true, color: COLORS.secondary,
  align: 'center', fontFace: 'Microsoft YaHei'
});

slide3.addText('7大类别105条知识\n语义检索\n向量匹配', {
  x: 0.6, y: boxY + 0.6, w: boxW - 0.2, h: 1.1,
  fontSize: 12, color: COLORS.textDark,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// LLM
slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 3.6, y: boxY, w: boxW, h: boxH,
  fill: { color: 'FFF3E0' },
  line: { color: COLORS.accent, width: 2 },
  rectRadius: 0.08
});

slide3.addText('💬 大语言模型', {
  x: 3.6, y: boxY + 0.15, w: boxW, h: 0.4,
  fontSize: 16, bold: true, color: 'E65100',
  align: 'center', fontFace: 'Microsoft YaHei'
});

slide3.addText('理解学生问题\n生成准确回答\n支持多轮对话', {
  x: 3.7, y: boxY + 0.6, w: boxW - 0.2, h: 1.1,
  fontSize: 12, color: COLORS.textDark,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// 微信公众号
slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 6.7, y: boxY, w: boxW, h: boxH,
  fill: { color: 'E8F5E9' },
  line: { color: '4CAF50', width: 2 },
  rectRadius: 0.08
});

slide3.addText('📢 微信公众号', {
  x: 6.7, y: boxY + 0.15, w: boxW, h: 0.4,
  fontSize: 16, bold: true, color: '2E7D32',
  align: 'center', fontFace: 'Microsoft YaHei'
});

slide3.addText('无需下载APP\n微信内即可使用\n覆盖率高', {
  x: 6.8, y: boxY + 0.6, w: boxW - 0.2, h: 1.1,
  fontSize: 12, color: COLORS.textDark,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// 工作流程
slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 3.6, w: 9, h: 1.5,
  fill: { color: COLORS.white },
  shadow: makeShadow(), rectRadius: 0.08
});

slide3.addText('🔄 工作流程', {
  x: 0.7, y: 3.75, w: 8.6, h: 0.35,
  fontSize: 16, bold: true, color: COLORS.primary,
  fontFace: 'Microsoft YaHei'
});

// 流程步骤
const steps = ['学生提问', '微信公众号', 'LLM推理', 'RAG检索', '智能回答'];
const stepW = 1.6;
const startX = 0.7;

steps.forEach((step, i) => {
  slide3.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: startX + i * (stepW + 0.1), y: 4.2, w: stepW, h: 0.6,
    fill: { color: i === 4 ? COLORS.primary : COLORS.secondary },
    rectRadius: 0.05
  });
  
  slide3.addText(step, {
    x: startX + i * (stepW + 0.1), y: 4.2, w: stepW, h: 0.6,
    fontSize: 12, color: COLORS.white,
    align: 'center', valign: 'middle',
    fontFace: 'Microsoft YaHei'
  });

  if (i < 4) {
    slide3.addText('→', {
      x: startX + i * (stepW + 0.1) + stepW, y: 4.2, w: 0.2, h: 0.6,
      fontSize: 16, color: COLORS.textLight,
      align: 'center', valign: 'middle'
    });
  }
});

// ========== 第4页：知识库内容 ==========
let slide4 = pptx.addSlide();
slide4.background = { color: COLORS.lightGray };

slide4.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide4.addText('📚 知识库内容', {
  x: 0.5, y: 0.15, w: 9, h: 0.6,
  fontSize: 28, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

// 7大类别 - 2行布局
const categories = [
  { emoji: '🏅', title: '奖助学金', desc: '国家奖学金、助学金、勤工俭学、助学贷款' },
  { emoji: '🏠', title: '宿舍管理', desc: '宿舍分配、调换申请、卫生检查、违规处理' },
  { emoji: '📝', title: '请假制度', desc: '请假类型、审批流程、证明材料、补办规定' },
  { emoji: '🔥', title: '消三比', desc: '学业预警、心理排查、违纪处理等三比内容' },
  { emoji: '📚', title: '教务相关', desc: '选课、转专业、休学复学、考试安排、成绩查询' },
  { emoji: '💼', title: '就业指导', desc: '简历制作、面试技巧、招聘信息、签约流程' },
  { emoji: '❓', title: 'FAQ', desc: '常见问题汇总，涵盖校园生活各方面' }
];

const catBoxW = 4.4;
const catBoxH = 1.1;
const catStartY = 1.1;

categories.forEach((cat, i) => {
  const row = Math.floor(i / 2);
  const col = i % 2;
  const x = 0.4 + col * 4.8;
  const y = catStartY + row * (catBoxH + 0.15);

  slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: catBoxW, h: catBoxH,
    fill: { color: COLORS.white },
    shadow: makeShadow(), rectRadius: 0.06
  });

  slide4.addText(`${cat.emoji} ${cat.title}`, {
    x: x + 0.15, y: y + 0.1, w: catBoxW - 0.3, h: 0.35,
    fontSize: 14, bold: true, color: COLORS.primary,
    fontFace: 'Microsoft YaHei'
  });

  slide4.addText(cat.desc, {
    x: x + 0.15, y: y + 0.5, w: catBoxW - 0.3, h: 0.5,
    fontSize: 11, color: COLORS.textLight,
    fontFace: 'Microsoft YaHei'
  });
});

// ========== 第5页：核心功能 ==========
let slide5 = pptx.addSlide();
slide5.background = { color: COLORS.lightGray };

slide5.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide5.addText('⚡ 核心功能', {
  x: 0.5, y: 0.15, w: 9, h: 0.6,
  fontSize: 28, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

const features = [
  { emoji: '💬', title: '即时问答', desc: '学生输入问题，3秒内返回答案，支持模糊查询' },
  { emoji: '🔍', title: '智能纠错', desc: '自动识别错别字和表达不清的问题，引导正确提问' },
  { emoji: '📎', title: '政策推送', desc: '主动推送最新通知（奖学金申请、假期安排等）' },
  { emoji: '📝', title: '转人工', desc: '复杂问题无缝转接辅导员，确保服务闭环' },
  { emoji: '📊', title: '数据统计', desc: '记录高频问题，为辅导员提供决策参考' }
];

features.forEach((feat, i) => {
  const y = 1.15 + i * 1.0;

  slide5.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: y, w: 9, h: 0.85,
    fill: { color: COLORS.white },
    shadow: makeShadow(), rectRadius: 0.06
  });

  // 左侧emoji圆圈
  slide5.addShape(pptx.shapes.OVAL, {
    x: 0.7, y: y + 0.15, w: 0.55, h: 0.55,
    fill: { color: COLORS.secondary }
  });

  slide5.addText(feat.emoji, {
    x: 0.7, y: y + 0.15, w: 0.55, h: 0.55,
    fontSize: 20, align: 'center', valign: 'middle'
  });

  slide5.addText(feat.title, {
    x: 1.4, y: y + 0.12, w: 7.9, h: 0.35,
    fontSize: 16, bold: true, color: COLORS.textDark,
    fontFace: 'Microsoft YaHei'
  });

  slide5.addText(feat.desc, {
    x: 1.4, y: y + 0.48, w: 7.9, h: 0.3,
    fontSize: 12, color: COLORS.textLight,
    fontFace: 'Microsoft YaHei'
  });
});

// ========== 第6页：预期效果 ==========
let slide6 = pptx.addSlide();
slide6.background = { color: COLORS.lightGray };

slide6.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide6.addText('📈 预期效果', {
  x: 0.5, y: 0.15, w: 9, h: 0.6,
  fontSize: 28, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

const effects = [
  { metric: '60%+', label: '人力解放', desc: '减少辅导员日常咨询工作量' },
  { metric: '95%+', label: '满意度提升', desc: '学生满意度预计提升至95%以上' },
  { metric: '7×24', label: '全天候服务', desc: '随时随地即时响应学生问题' },
  { metric: '99.9%', label: '服务可用性', desc: '多副本部署，自动故障转移' }
];

effects.forEach((eff, i) => {
  const x = 0.4 + (i % 2) * 4.8;
  const y = 1.1 + Math.floor(i / 2) * 2.2;

  slide6.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: y, w: 4.4, h: 2.0,
    fill: { color: COLORS.white },
    shadow: makeShadow(), rectRadius: 0.08
  });

  slide6.addText(eff.metric, {
    x: x, y: y + 0.2, w: 4.4, h: 0.7,
    fontSize: 36, bold: true, color: COLORS.primary,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide6.addText(eff.label, {
    x: x, y: y + 0.95, w: 4.4, h: 0.4,
    fontSize: 16, bold: true, color: COLORS.textDark,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide6.addText(eff.desc, {
    x: x + 0.2, y: y + 1.4, w: 4, h: 0.4,
    fontSize: 11, color: COLORS.textLight,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
});

// ========== 第7页：技术亮点 ==========
let slide7 = pptx.addSlide();
slide7.background = { color: COLORS.lightGray };

slide7.addShape(pptx.shapes.RECTANGLE, {
  x: 0, y: 0, w: 10, h: 0.9,
  fill: { color: COLORS.primary }
});

slide7.addText('🔧 技术亮点', {
  x: 0.5, y: 0.15, w: 9, h: 0.6,
  fontSize: 28, bold: true, color: COLORS.white,
  fontFace: 'Microsoft YaHei', margin: 0
});

const techPoints = [
  { emoji: '🔍', title: 'RAG检索增强', desc: '结合向量检索与关键词匹配，答案准确率行业领先' },
  { emoji: '🧩', title: '知识库管理', desc: '可视化知识库配置，支持一键更新和版本管理' },
  { emoji: '🛡️', title: '内容安全', desc: '内置敏感词过滤和回答审核机制，确保输出合规' },
  { emoji: '🔄', title: '容灾备份', desc: '多副本部署，自动故障转移，服务可用性99.9%' },
  { emoji: '📱', title: '微信生态', desc: '无缝对接微信公众号，无需学生额外安装软件' }
];

techPoints.forEach((tp, i) => {
  const y = 1.15 + i * 1.0;

  slide7.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: y, w: 9, h: 0.85,
    fill: { color: COLORS.white },
    shadow: makeShadow(), rectRadius: 0.06
  });

  slide7.addShape(pptx.shapes.OVAL, {
    x: 0.7, y: y + 0.15, w: 0.55, h: 0.55,
    fill: { color: COLORS.accent }
  });

  slide7.addText(tp.emoji, {
    x: 0.7, y: y + 0.15, w: 0.55, h: 0.55,
    fontSize: 20, align: 'center', valign: 'middle'
  });

  slide7.addText(tp.title, {
    x: 1.4, y: y + 0.12, w: 7.9, h: 0.35,
    fontSize: 16, bold: true, color: COLORS.textDark,
    fontFace: 'Microsoft YaHei'
  });

  slide7.addText(tp.desc, {
    x: 1.4, y: y + 0.48, w: 7.9, h: 0.3,
    fontSize: 12, color: COLORS.textLight,
    fontFace: 'Microsoft YaHei'
  });
});

// ========== 第8页：下一步计划 ==========
let slide8 = pptx.addSlide();
slide8.background = { color: COLORS.primary };

slide8.addText('🚀 下一步计划', {
  x: 0.5, y: 0.4, w: 9, h: 0.8,
  fontSize: 32, bold: true, color: COLORS.white,
  align: 'center', fontFace: 'Microsoft YaHei'
});

// 三阶段卡片
const phases = [
  { phase: '第一阶段', time: '1个月', title: '知识库建设', desc: '完成知识库建设和内部测试' },
  { phase: '第二阶段', time: '2个月', title: '小范围试点', desc: '小范围试点，收集反馈优化体验' },
  { phase: '第三阶段', time: '3个月', title: '全量上线', desc: '全量上线，持续迭代知识库' }
];

phases.forEach((ph, i) => {
  const x = 0.5 + i * 3.1;

  slide8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 1.5, w: 2.9, h: 2.8,
    fill: { color: COLORS.white },
    shadow: makeShadow(), rectRadius: 0.1
  });

  // 阶段标签
  slide8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: x + 0.2, y: 1.7, w: 2.5, h: 0.4,
    fill: { color: COLORS.accent },
    rectRadius: 0.05
  });

  slide8.addText(ph.phase, {
    x: x + 0.2, y: 1.7, w: 2.5, h: 0.4,
    fontSize: 14, bold: true, color: COLORS.white,
    align: 'center', valign: 'middle',
    fontFace: 'Microsoft YaHei'
  });

  slide8.addText(ph.time, {
    x: x, y: 2.25, w: 2.9, h: 0.4,
    fontSize: 20, bold: true, color: COLORS.primary,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide8.addText(ph.title, {
    x: x + 0.15, y: 2.75, w: 2.6, h: 0.4,
    fontSize: 14, bold: true, color: COLORS.textDark,
    align: 'center', fontFace: 'Microsoft YaHei'
  });

  slide8.addText(ph.desc, {
    x: x + 0.15, y: 3.2, w: 2.6, h: 0.8,
    fontSize: 11, color: COLORS.textLight,
    align: 'center', fontFace: 'Microsoft YaHei'
  });
});

// 合作模式
slide8.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
  x: 0.5, y: 4.5, w: 9, h: 0.8,
  fill: { color: '1A4073' },
  rectRadius: 0.06
});

slide8.addText('🤝 合作模式：技术支持 + 知识库共建 + 持续运营服务', {
  x: 0.5, y: 4.5, w: 9, h: 0.8,
  fontSize: 16, color: COLORS.white,
  align: 'center', valign: 'middle',
  fontFace: 'Microsoft YaHei'
});

// 保存文件
pptx.writeFile({ fileName: '学院AI辅导员系统介绍.pptx' })
  .then(() => {
    console.log('✅ PPT生成成功: 学院AI辅导员系统介绍.pptx');
  })
  .catch(err => {
    console.error('❌ 生成失败:', err);
  });
