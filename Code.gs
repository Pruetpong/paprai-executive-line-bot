// Code.gs - Main Webhook Handler for Academic Affairs AI Assistant
// โค้ดหลักสำหรับจัดการ webhook และการสนทนา

/**
 * ====================================
 * WEBHOOK HANDLER
 * ====================================
 */

function doPost(e) {
  const startTime = new Date();
  console.log(`🌐 Webhook received at ${startTime.toISOString()}`);
  
  try {
    if (!e.postData || !e.postData.contents) {
      console.error('❌ Invalid request format');
      return createResponse('Invalid request', 400);
    }

    const contents = JSON.parse(e.postData.contents);
    if (!contents.events || !Array.isArray(contents.events)) {
      console.error('❌ Invalid events format');
      return createResponse('Invalid events', 400);
    }

    console.log(`📨 Processing ${contents.events.length} event(s)`);
    
    const credentials = getCredentials();
    if (!validateCredentials(credentials)) {
      console.error('❌ Missing required credentials');
      return createResponse('Service unavailable', 503);
    }

    for (const event of contents.events) {
      console.log(`🔄 Processing event type: ${event.type}`);
      processEvent(event, credentials);
    }

    const processingTime = new Date() - startTime;
    console.log(`✅ Webhook completed in ${processingTime}ms`);

    return createResponse({ status: 'success', processed: contents.events.length }, 200);

  } catch (error) {
    console.error('❌ Webhook error:', error);
    return createResponse('Internal error', 500);
  }
}

/**
 * ====================================
 * EVENT PROCESSING
 * ====================================
 */

function processEvent(event, credentials) {
  try {
    const { type, replyToken, source } = event;
    const userId = source?.userId;

    // ตรวจสอบ Access Control
    const userProfile = getUserProfile(userId);
    if (!userProfile) {
      console.log(`🚫 Unauthorized user: ${userId}`);
      handleUnauthorizedUser(replyToken);
      return;
    }

    // เริ่ม Loading Indicator
    if (userId && ['message', 'follow'].includes(type)) {
      startLoading(userId);
    }

    // จัดการ Event แต่ละประเภท
    switch (type) {
      case 'message':
        handleMessage(event, userProfile);
        break;
      case 'follow':
        handleFollow(replyToken, userProfile);
        break;
      case 'unfollow':
        handleUnfollow(userId);
        break;
      default:
        console.log(`⚠️ Unhandled event type: ${type}`);
    }
    
  } catch (error) {
    console.error(`❌ Error processing event:`, error);
    
    if (event.replyToken) {
      try {
        sendTextMessage(event.replyToken, "ขอภัยค่ะ เกิดข้อผิดพลาดในการประมวลผล กรุณาลองใหม่อีกครั้ง 🙏");
      } catch (replyError) {
        console.error('❌ Failed to send error message:', replyError);
      }
    }
  }
}

/**
 * ====================================
 * MESSAGE HANDLING
 * ====================================
 */

function handleMessage(event, userProfile) {
  const { message, replyToken } = event;
  const { type: messageType } = message;
  const quoteToken = message.quoteToken || null;

  console.log(`💬 Message type: ${messageType} from ${userProfile.name}`);

  switch (messageType) {
    case 'text':
      handleTextMessage(event, userProfile, quoteToken);
      break;
    case 'image':
      handleImageMessage(event, userProfile, quoteToken);
      break;
    default:
      console.log(`⚠️ Unsupported message type: ${messageType}`);
      sendTextMessage(replyToken, "ขอภัยค่ะ ตอนนี้รองรับเฉพาะข้อความและรูปภาพเท่านั้น 📝📷");
  }
}

function handleTextMessage(event, userProfile, quoteToken) {
  const { message, replyToken } = event;
  const userMessage = message.text.trim();

  console.log(`📝 Text: "${userMessage}" from ${userProfile.name}`);

  // ตรวจสอบคำสั่งพิเศษ
  if (handleSpecialCommands(userMessage, replyToken, userProfile, quoteToken)) {
    return;
  }

  // ประมวลผลข้อความปกติด้วย AI
  processAIMessage(userMessage, replyToken, userProfile, quoteToken);
}

function handleImageMessage(event, userProfile, quoteToken) {
  const { message, replyToken } = event;
  const messageId = message.id;
  
  console.log(`📷 Image received from ${userProfile.name}`);

  try {
    // ดาวน์โหลดรูปภาพจาก LINE
    const imageBlob = getImageFromLine(messageId);
    
    // แปลงเป็น Base64
    const base64Image = Utilities.base64Encode(imageBlob.getBytes());
    
    // ดึงคำอธิบายจาก caption (ถ้ามี)
    const caption = message.text || "กรุณาวิเคราะห์รูปภาพนี้และให้คำแนะนำที่เหมาะสม";
    
    // วิเคราะห์รูปภาพด้วย OpenAI Vision
    processImageAnalysis(base64Image, caption, replyToken, userProfile, quoteToken);
    
  } catch (error) {
    console.error('❌ Image processing error:', error);
    sendTextMessage(replyToken, "ขอภัยค่ะ เกิดข้อผิดพลาดในการประมวลผลรูปภาพ กรุณาลองใหม่อีกครั้ง 🙏", quoteToken);
  }
}

/**
 * ====================================
 * SPECIAL COMMANDS
 * ====================================
 */

function handleSpecialCommands(userMessage, replyToken, userProfile, quoteToken) {
  const command = userMessage.toLowerCase();
  
  switch (command) {
    case '/clear':
    case 'clear':
      clearChatHistory(userProfile.userId);
      sendTextMessage(replyToken, "✅ ลบประวัติการสนทนาเรียบร้อยแล้วค่ะ", quoteToken);
      return true;
      
    case '/help':
    case 'help':
      sendHelpMessage(replyToken, userProfile, quoteToken);
      return true;
      
    case '/status':
    case 'status':
      sendStatusMessage(replyToken, userProfile, quoteToken);
      return true;
      
    case '/profile':
    case 'profile':
      sendProfileMessage(replyToken, userProfile, quoteToken);
      return true;
      
    case '/analytics':
    case 'analytics':
      sendAnalyticsMessage(replyToken, userProfile, quoteToken);
      return true;
      
    case '/tags':
    case 'tags':
      sendTagsMessage(replyToken, userProfile, quoteToken);
      return true;
      
    case '/export':
    case 'export':
      exportChatHistory(replyToken, userProfile, quoteToken);
      return true;
      
    default:
      return false;
  }
}

/**
 * ====================================
 * AI MESSAGE PROCESSING
 * ====================================
 */

function processAIMessage(userMessage, replyToken, userProfile, quoteToken) {
  try {
    console.log(`🤖 Processing AI request for ${userProfile.name}`);
    
    // ดึงประวัติการสนทนา
    const chatHistory = getChatHistory(userProfile.userId);
    
    // สร้าง Prompt พร้อม Context
    const systemPrompt = getSystemPrompt(userProfile);
    const userPrompt = constructUserPrompt(chatHistory, userMessage, userProfile);
    
    // เรียก OpenAI API
    const startAI = new Date();
    const response = generateAIResponse(systemPrompt, userPrompt);
    const aiTime = new Date() - startAI;
    
    // Auto-tag หมวดหมู่
    const category = autoTagMessage(userMessage);
    
    // บันทึกประวัติ
    saveChatHistory(userProfile, userMessage, response, 'text', null, category, aiTime);
    
    // ส่งคำตอบกลับ
    sendTextMessage(replyToken, response, quoteToken);
    
    console.log(`✅ AI response sent (${aiTime}ms)`);
    
  } catch (error) {
    console.error('❌ AI processing error:', error);
    
    let errorMessage = "ป้าไพรขอภัยค่ะ เกิดข้อผิดพลาดในการประมวลผล กรุณาลองใหม่อีกครั้งนะคะ 🙏";
    
    if (error.message.includes('API key')) {
      errorMessage = "ป้าไพรขอภัยค่ะ มีปัญหาในการเชื่อมต่อระบบ AI กรุณาติดต่อผู้ดูแลระบบนะคะ 🔧";
    } else if (error.message.includes('rate limit')) {
      errorMessage = "ป้าไพรขอภัยค่ะ ระบบใช้งานหนักเกินไป กรุณารอสักครู่แล้วลองใหม่นะคะ ⏳";
    }
    
    sendTextMessage(replyToken, errorMessage, quoteToken);
  }
}

function processImageAnalysis(base64Image, caption, replyToken, userProfile, quoteToken) {
  try {
    console.log(`🖼️ Analyzing image for ${userProfile.name}`);
    
    // ดึงประวัติการสนทนา
    const chatHistory = getChatHistory(userProfile.userId);
    
    // สร้าง System Prompt
    const systemPrompt = getSystemPrompt(userProfile);
    
    // เรียก OpenAI Vision API
    const startAI = new Date();
    const response = analyzeImage(systemPrompt, base64Image, caption, chatHistory, userProfile);
    const aiTime = new Date() - startAI;
    
    // Auto-tag
    const category = autoTagMessage(caption);
    
    // สร้าง placeholder image URL (เนื่องจากไม่เก็บรูปใน Drive)
    const imageUrl = `[Image analyzed at ${new Date().toISOString()}]`;
    
    // บันทึกประวัติ
    saveChatHistory(userProfile, caption, response, 'image', imageUrl, category, aiTime);
    
    // ส่งคำตอบกลับ
    sendTextMessage(replyToken, response, quoteToken);
    
    console.log(`✅ Image analysis sent (${aiTime}ms)`);
    
  } catch (error) {
    console.error('❌ Image analysis error:', error);
    sendTextMessage(replyToken, "ขอภัยค่ะ เกิดข้อผิดพลาดในการวิเคราะห์รูปภาพ กรุณาลองใหม่อีกครั้ง 🙏", quoteToken);
  }
}

/**
 * ====================================
 * CHAT HISTORY MANAGEMENT
 * ====================================
 */

function getChatHistory(userId) {
  try {
    const credentials = getCredentials();
    if (!credentials.SPREADSHEET_ID) {
      console.warn('⚠️ No spreadsheet configured');
      return [];
    }

    const ss = SpreadsheetApp.openById(credentials.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Chat History');
    
    if (!sheet) {
      console.warn('⚠️ Chat History sheet not found');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    
    // กรองเฉพาะข้อความของ user นี้ และเอาล่าสุด MAX_HISTORY รายการ
    const userHistory = data
      .filter(row => row[1] === userId) // Column B = User ID
      .slice(-APP_CONFIG.MAX_HISTORY)
      .map(row => ({
        userMessage: row[5],    // Column F = User Message
        aiResponse: row[6]      // Column G = AI Response
      }));
      
    console.log(`📚 Retrieved ${userHistory.length} history entries`);
    return userHistory;
    
  } catch (error) {
    console.error('❌ Error retrieving chat history:', error);
    return [];
  }
}

function saveChatHistory(userProfile, userMessage, aiResponse, messageType, imageUrl, category, responseTime) {
  try {
    const credentials = getCredentials();
    if (!credentials.SPREADSHEET_ID) {
      console.warn('⚠️ No spreadsheet configured');
      return;
    }

    const ss = SpreadsheetApp.openById(credentials.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Chat History');
    
    if (!sheet) {
      console.warn('⚠️ Chat History sheet not found');
      return;
    }
    
    const timestamp = new Date();
    const tokensUsed = estimateTokens(userMessage + aiResponse);
    
    // เพิ่มแถวใหม่: Timestamp, User ID, User Name, Role, Message Type, User Message, AI Response, Tokens, Response Time, Auto Tags, Image URL
    sheet.appendRow([
      timestamp,
      userProfile.userId,
      userProfile.name,
      userProfile.role,
      messageType,
      userMessage,
      aiResponse,
      tokensUsed,
      responseTime,
      category,
      imageUrl || ''
    ]);
    
    console.log('💾 Chat history saved');
    
    // อัปเดต Analytics
    updateAnalytics(userProfile, category, tokensUsed, responseTime);
    
  } catch (error) {
    console.error('❌ Error saving chat history:', error);
  }
}

function clearChatHistory(userId) {
  try {
    const credentials = getCredentials();
    if (!credentials.SPREADSHEET_ID) {
      console.warn('⚠️ No spreadsheet configured');
      return;
    }

    const ss = SpreadsheetApp.openById(credentials.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Chat History');
    
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    
    // หาแถวที่ต้องลบ (เริ่มจากท้ายขึ้นบน)
    for (let i = data.length - 1; i >= 1; i--) { // เริ่มจาก index 1 เพื่อข้าม header
      if (data[i][1] === userId) { // Column B = User ID
        rowsToDelete.push(i + 1); // +1 เพราะ sheet index เริ่มจาก 1
      }
    }
    
    // ลบแถว (จากท้ายขึ้นบนเพื่อไม่ให้ index เปลี่ยน)
    rowsToDelete.forEach(rowNum => {
      sheet.deleteRow(rowNum);
    });
    
    console.log(`🗑️ Cleared ${rowsToDelete.length} history entries`);
    
  } catch (error) {
    console.error('❌ Error clearing chat history:', error);
  }
}

/**
 * ====================================
 * COMMAND RESPONSES
 * ====================================
 */

function sendHelpMessage(replyToken, userProfile, quoteToken) {
  const helpText = `🤖 คู่มือการใช้งาน ป้าไพร

สวัสดีค่ะ ${userProfile.name} 
ป้าไพรยินดีช่วยเหลือคุณในทุกเรื่องค่ะ 😊

📋 คำสั่งที่ใช้ได้
━━━━━━━━━━━━━━━━━━━━━━

🗑️ /clear
   ลบประวัติการสนทนาทั้งหมด

📊 /status
   ดูสถานะระบบและการใช้งาน

👤 /profile
   ดูข้อมูลความสามารถของป้าไพร

📈 /analytics
   ดูสถิติการใช้งานของคุณ

🏷️ /tags
   ดูหมวดหมู่คำถามที่ถามบ่อย

📤 /export
   ส่งออกประวัติการสนทนา

❓ /help
   แสดงคู่มือนี้

💡 การใช้งานทั่วไป
━━━━━━━━━━━━━━━━━━━━━━

📝 พิมพ์ข้อความธรรมดา
   ปรึกษาปัญหาหรือถามคำถาม

📷 ส่งรูปภาพ
   วิเคราะห์เอกสาร แผนผัง ห้องเรียน

🎯 ป้าไพรพร้อมช่วยเหลือด้าน
   หลักสูตร การสอน นวัตกรรม
   การประเมินผล พัฒนาครู STEAM
   เทคโนโลยีการศึกษา และอื่นๆ

มีอะไรให้ป้าไพรช่วยไหมคะ ${userProfile.name} 😊`;

  sendTextMessage(replyToken, helpText, quoteToken);
}

function sendStatusMessage(replyToken, userProfile, quoteToken) {
  try {
    const chatHistory = getChatHistory(userProfile.userId);
    const analytics = getUserAnalytics(userProfile.userId);
    
    const statusText = `📊 สถานะระบบ

👤 ผู้ใช้งาน ${userProfile.name}
   ตำแหน่ง ${getRoleDisplayName(userProfile.role)}

💬 ประวัติการสนทนา ${chatHistory.length}/${APP_CONFIG.MAX_HISTORY} ข้อความ

📈 สถิติการใช้งาน
   จำนวนข้อความ ${analytics.totalMessages || 0} ข้อความ
   หมวดหมู่ยอดนิยม ${analytics.topCategory || 'ยังไม่มีข้อมูล'}
   เวลาตอบเฉลี่ย ${analytics.avgResponseTime || 'N/A'} ms

🤖 AI Model ${APP_CONFIG.AI_MODEL}
⏰ เวลา ${new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' })}

ระบบทำงานปกติค่ะ ✅
มีอะไรให้ป้าไพรช่วยไหมคะ`;

    sendTextMessage(replyToken, statusText, quoteToken);
    
  } catch (error) {
    console.error('Error getting status:', error);
    sendTextMessage(replyToken, "ป้าไพรไม่สามารถดึงข้อมูลสถานะได้ค่ะ กรุณาลองใหม่อีกครั้งนะคะ 🙏", quoteToken);
  }
}

function sendProfileMessage(replyToken, userProfile, quoteToken) {
  const roleInfo = getRoleInfo(userProfile.role);
  
  const profileText = `👤 ข้อมูลของป้าไพร

สวัสดีค่ะ ${userProfile.name} 🙏
ป้าไพรยินดีที่จะช่วยเหลือคุณค่ะ

🤖 เกี่ยวกับป้าไพร
━━━━━━━━━━━━━━━━━━━━━━

ชื่อ ป้าไพร (PAPRAI)
บทบาท ผู้ช่วย AI ที่ปรึกษาวิชาการ
สำหรับ ${roleInfo.displayName}

💡 ความเชี่ยวชาญที่ป้าไพรมี
━━━━━━━━━━━━━━━━━━━━━━

${roleInfo.expertise}

🎯 รูปแบบการให้คำแนะนำ
━━━━━━━━━━━━━━━━━━━━━━

${roleInfo.approach}

📚 คำถามที่เหมาะถามป้าไพร
━━━━━━━━━━━━━━━━━━━━━━

${roleInfo.sampleQuestions}

มีอะไรให้ป้าไพรช่วยไหมคะ 😊`;

  sendTextMessage(replyToken, profileText, quoteToken);
}

function sendAnalyticsMessage(replyToken, userProfile, quoteToken) {
  try {
    const analytics = getUserAnalytics(userProfile.userId);
    const topCategories = analytics.categories || {};
    
    let categoriesText = '';
    Object.entries(topCategories)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .forEach(([cat, count]) => {
        categoriesText += `   ${cat} (${count} ครั้ง)\n`;
      });
    
    const analyticsText = `📈 สถิติการใช้งานของ ${userProfile.name}

📊 สรุปภาพรวม
━━━━━━━━━━━━━━━━━━━━━━

💬 จำนวนข้อความทั้งหมด
   ${analytics.totalMessages || 0} ข้อความ

🏷️ หมวดหมู่ที่ปรึกษาบ่อยสุด
${categoriesText || '   ยังไม่มีข้อมูล'}

⏱️ เวลาตอบกลับเฉลี่ย
   ${analytics.avgResponseTime || 'N/A'} ms

📅 ช่วงเวลาการใช้งาน
   ${analytics.dateRange || 'ยังไม่มีข้อมูล'}

💾 ข้อมูล ณ วันที่
   ${new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' })}

มีอะไรให้ป้าไพรช่วยเพิ่มเติมไหมคะ 🙏`;

    sendTextMessage(replyToken, analyticsText, quoteToken);
    
  } catch (error) {
    console.error('Error getting analytics:', error);
    sendTextMessage(replyToken, "ป้าไพรไม่สามารถดึงข้อมูลสถิติได้ค่ะ กรุณาลองใหม่อีกครั้งนะคะ 🙏", quoteToken);
  }
}

function sendTagsMessage(replyToken, userProfile, quoteToken) {
  const tagsText = `🏷️ หมวดหมู่คำถามที่ป้าไพรรองรับ

📚 หมวดหมู่ทั้งหมด
━━━━━━━━━━━━━━━━━━━━━━

${APP_CONFIG.CATEGORIES.map((cat, i) => `${i + 1}. ${cat}`).join('\n')}

💡 ป้าไพรจะจัดหมวดหมู่คำถาม
   ของคุณโดยอัตโนมัติ เพื่อวิเคราะห์
   และปรับปรุงคุณภาพการให้คำปรึกษา

📊 ดูสถิติการใช้งานได้ที่
   /analytics

ถามป้าไพรได้ทุกหมวดหมู่เลยนะคะ 😊`;

  sendTextMessage(replyToken, tagsText, quoteToken);
}

function exportChatHistory(replyToken, userProfile, quoteToken) {
  try {
    const credentials = getCredentials();
    if (!credentials.SPREADSHEET_ID) {
      sendTextMessage(replyToken, "ยังไม่ได้ตั้งค่า Spreadsheet ค่ะ 🙏", quoteToken);
      return;
    }

    const ss = SpreadsheetApp.openById(credentials.SPREADSHEET_ID);
    const url = ss.getUrl();
    
    const exportText = `📤 ส่งออกประวัติการสนทนา

คุณสามารถดูประวัติการสนทนาทั้งหมดได้ที่

🔗 ${url}

📊 ข้อมูลใน Spreadsheet ประกอบด้วย
   • Chat History (ประวัติการสนทนา)
   • Usage Analytics (สถิติการใช้งาน)
   • Weekly Reports (รายงานประจำสัปดาห์)

🔐 หมายเหตุ เฉพาะผู้ที่มีสิทธิ์เท่านั้น
   จึงจะเข้าถึงข้อมูลได้

ขอบคุณที่ใช้บริการค่ะ 🙏`;

    sendTextMessage(replyToken, exportText, quoteToken);
    
  } catch (error) {
    console.error('Error exporting history:', error);
    sendTextMessage(replyToken, "ไม่สามารถส่งออกข้อมูลได้ค่ะ 🙏", quoteToken);
  }
}

/**
 * ====================================
 * EVENT HANDLERS
 * ====================================
 */

function handleFollow(replyToken, userProfile) {
  console.log(`👋 User ${userProfile.name} followed the bot`);
  
  const welcomeText = `สวัสดีค่ะ ${userProfile.name} 🙏

ฉันคือ ป้าไพร (PAPRAI)
AI Assistant สำหรับฝ่ายวิชาการและนวัตกรรมการเรียนรู้
โรงเรียนสาธิต มหาวิทยาลัยศิลปากร (มัธยมศึกษา)

🤖 ป้าไพรพร้อมให้คำปรึกษา
   และสนับสนุนการทำงาน
   ด้านวิชาการของคุณค่ะ

💡 เริ่มต้นใช้งาน
   พิมพ์ /help เพื่อดูคู่มือ

มีอะไรให้ป้าไพรช่วยไหมคะ 😊`;

  sendTextMessage(replyToken, welcomeText);
}

function handleUnfollow(userId) {
  console.log(`👋 User ${userId} unfollowed the bot`);
  // สามารถเพิ่มการจัดการเพิ่มเติมได้ เช่น การทำความสะอาดข้อมูล
}

function handleUnauthorizedUser(replyToken) {
  const message = `สวัสดีค่ะ 🙏

ป้าไพรขออภัยด้วยนะคะ
ระบบนี้เป็นบริการเฉพาะสำหรับ
ผู้บริหารฝ่ายวิชาการและนวัตกรรมการเรียนรู้เท่านั้นค่ะ

หากคุณต้องการใช้งาน
กรุณาติดต่อผู้ดูแลระบบ
เพื่อขอสิทธิ์การเข้าถึงนะคะ

ขอบคุณที่ให้ความสนใจค่ะ 🙏`;

  sendTextMessage(replyToken, message);
}

/**
 * ====================================
 * LINE API FUNCTIONS
 * ====================================
 */

function sendTextMessage(replyToken, text, quoteToken = null) {
  try {
    const message = { type: 'text', text: text };
    if (quoteToken) {
      message.quoteToken = quoteToken;
    }
    
    replyMessage(replyToken, [message]);
    console.log('💬 Text message sent');
    
  } catch (error) {
    console.error('❌ Error sending text message:', error);
    throw error;
  }
}

function replyMessage(replyToken, messages) {
  const credentials = getCredentials();
  if (!credentials.LINE_CHANNEL_ACCESS_TOKEN) {
    throw new Error('LINE_CHANNEL_ACCESS_TOKEN not configured');
  }

  const url = 'https://api.line.me/v2/bot/message/reply';
  const options = {
    headers: {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': `Bearer ${credentials.LINE_CHANNEL_ACCESS_TOKEN}`,
    },
    method: 'post',
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: messages
    }),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode !== 200) {
    const errorText = response.getContentText();
    console.error(`LINE API Error (${responseCode}):`, errorText);
    throw new Error(`LINE API returned status ${responseCode}`);
  }
}

function pushMessage(userId, messages) {
  const credentials = getCredentials();
  if (!credentials.LINE_CHANNEL_ACCESS_TOKEN) {
    throw new Error('LINE_CHANNEL_ACCESS_TOKEN not configured');
  }

  const url = 'https://api.line.me/v2/bot/message/push';
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${credentials.LINE_CHANNEL_ACCESS_TOKEN}`
    },
    payload: JSON.stringify({
      to: userId,
      messages: messages
    }),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode !== 200) {
    const errorText = response.getContentText();
    console.error(`LINE Push API Error (${responseCode}):`, errorText);
    throw new Error(`LINE Push API returned status ${responseCode}`);
  }
}

function startLoading(userId) {
  try {
    const credentials = getCredentials();
    if (!credentials.LINE_CHANNEL_ACCESS_TOKEN || !userId) return;

    const url = 'https://api.line.me/v2/bot/chat/loading/start';
    const options = {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${credentials.LINE_CHANNEL_ACCESS_TOKEN}`,
      },
      method: 'post',
      payload: JSON.stringify({ chatId: userId }),
      muteHttpExceptions: true
    };

    UrlFetchApp.fetch(url, options);
    console.log('⏳ Loading indicator started');
    
  } catch (error) {
    console.error('⚠️ Failed to start loading indicator:', error);
  }
}

function getImageFromLine(messageId) {
  const credentials = getCredentials();
  if (!credentials.LINE_CHANNEL_ACCESS_TOKEN) {
    throw new Error('LINE_CHANNEL_ACCESS_TOKEN not configured');
  }

  const url = `https://api-data.line.me/v2/bot/message/${messageId}/content`;
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': `Bearer ${credentials.LINE_CHANNEL_ACCESS_TOKEN}` },
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    throw new Error(`Failed to download image: ${responseCode}`);
  }

  return response.getBlob();
}

/**
 * ====================================
 * UTILITY FUNCTIONS
 * ====================================
 */

function createResponse(data, statusCode = 200) {
  const output = typeof data === 'string' ? { message: data } : data;
  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON)
    .setStatusCode(200);
}

function estimateTokens(text) {
  // ประมาณการ tokens (ภาษาไทย ~1.5 tokens ต่อคำ)
  return Math.ceil(text.length / 3);
}

function validateCredentials(credentials) {
  const required = ['LINE_CHANNEL_ACCESS_TOKEN', 'OPENAI_API_KEY'];
  return required.every(key => credentials[key] && credentials[key].trim() !== '');
}
