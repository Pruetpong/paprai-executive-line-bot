// Config.gs - Configuration and AI Engine for Academic Affairs AI Assistant
// ไฟล์ Configuration และระบบ AI

/**
 * ====================================
 * APPLICATION CONFIGURATION
 * ====================================
 */

const APP_CONFIG = {
  // จำนวนประวัติการสนทนาสูงสุดที่เก็บไว้
  MAX_HISTORY: 10,
  
  // AI Model ที่ใช้
  AI_MODEL: 'gpt-4o',  // รองรับ Vision
  OPENAI_ENDPOINT: 'https://api.openai.com/v1/chat/completions',
  
  // พารามิเตอร์ AI
  AI_PARAMS: {
    temperature: 0.7,
    max_tokens: 3000
  },
  
  // หมวดหมู่คำถามที่รองรับ (สำหรับ Auto-tagging)
  CATEGORIES: [
    'หลักสูตร',
    'การสอน',
    'การประเมินผล',
    'นวัตกรรม',
    'เทคโนโลยี',
    'พัฒนาครู',
    'นโยบาย',
    'บริหารจัดการ',
    'ประกันคุณภาพ',
    'STEAM',
    'วิจัย',
    'อื่นๆ'
  ],
  
  // ตั้งค่าความปลอดภัย
  SECURITY: {
    REQUIRE_WHITELIST: true,
    RATE_LIMIT_PER_HOUR: 100
  }
};

/**
 * ====================================
 * CREDENTIALS MANAGEMENT
 * ====================================
 */

/**
 * ฟังก์ชันตั้งค่า Credentials ครั้งแรก
 * 
 * วิธีใช้งาน:
 * 1. เปิด Apps Script Editor
 * 2. ไปที่ Extensions > Apps Script
 * 3. เรียกใช้ฟังก์ชัน setupCredentials()
 */
function setupCredentials() {
  const credentials = {
    // กรอกข้อมูลของคุณที่นี่
    OPENAI_API_KEY: 'YOUR_OPENAI_API_KEY_HERE',
    LINE_CHANNEL_ACCESS_TOKEN: 'YOUR_LINE_CHANNEL_ACCESS_TOKEN_HERE',
    SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',
    
    // ข้อมูลผู้ใช้ทั้ง 2 ท่าน
    USER_PROFILES: {
      // ผู้ใช้ที่ 1: อาจารย์พรึด (ผู้ช่วยผู้อำนวยการ)
      'YOUR_USER_ID_1': {
        name: 'อาจารย์พรึด',
        role: 'ASSISTANT_DIRECTOR',
        department: 'Academic Affairs and Learning Innovation',
        active: true
      },
      // ผู้ใช้ที่ 2: อาจารย์ห้อย (รองผู้อำนวยการ)
      'YOUR_USER_ID_2': {
        name: 'อาจารย์ห้อย',
        role: 'DEPUTY_DIRECTOR',
        department: 'Academic Affairs and Learning Innovation',
        active: true
      }
    }
  };
  
  // บันทึก Credentials
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('OPENAI_API_KEY', credentials.OPENAI_API_KEY);
  scriptProperties.setProperty('LINE_CHANNEL_ACCESS_TOKEN', credentials.LINE_CHANNEL_ACCESS_TOKEN);
  scriptProperties.setProperty('SPREADSHEET_ID', credentials.SPREADSHEET_ID);
  scriptProperties.setProperty('USER_PROFILES', JSON.stringify(credentials.USER_PROFILES));
  
  console.log('✅ Credentials setup completed');
  console.log('กรุณาแทนที่ YOUR_OPENAI_API_KEY_HERE, YOUR_LINE_CHANNEL_ACCESS_TOKEN_HERE, YOUR_SPREADSHEET_ID_HERE และ YOUR_USER_ID_1, YOUR_USER_ID_2 ด้วยค่าจริง');
}

function getCredentials() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    
    return {
      OPENAI_API_KEY: scriptProperties.getProperty('OPENAI_API_KEY'),
      LINE_CHANNEL_ACCESS_TOKEN: scriptProperties.getProperty('LINE_CHANNEL_ACCESS_TOKEN'),
      SPREADSHEET_ID: scriptProperties.getProperty('SPREADSHEET_ID'),
      USER_PROFILES: JSON.parse(scriptProperties.getProperty('USER_PROFILES') || '{}')
    };
  } catch (error) {
    console.error('❌ Error getting credentials:', error);
    throw new Error('Failed to retrieve credentials');
  }
}

/**
 * ====================================
 * USER PROFILE MANAGEMENT
 * ====================================
 */

function getUserProfile(userId) {
  try {
    const credentials = getCredentials();
    const userProfiles = credentials.USER_PROFILES || {};
    
    const profile = userProfiles[userId];
    
    if (!profile || !profile.active) {
      return null; // ผู้ใช้ไม่ได้รับอนุญาต
    }
    
    return {
      userId: userId,
      name: profile.name,
      role: profile.role,
      department: profile.department
    };
    
  } catch (error) {
    console.error('Error getting user profile:', error);
    return null;
  }
}

function getRoleDisplayName(role) {
  const roleNames = {
    'DEPUTY_DIRECTOR': 'รองผู้อำนวยการฝ่ายวิชาการและนวัตกรรมการเรียนรู้',
    'ASSISTANT_DIRECTOR': 'ผู้ช่วยผู้อำนวยการด้านนวัตกรรมการเรียนรู้'
  };
  return roleNames[role] || 'ไม่ระบุ';
}

function getRoleInfo(role) {
  if (role === 'DEPUTY_DIRECTOR') {
    return {
      displayName: 'รองผู้อำนวยการฝ่ายวิชาการและนวัตกรรมการเรียนรู้',
      expertise: `📚 การบริหารจัดการวิชาการ
🎓 การประกันคุณภาพการศึกษา
📖 การพัฒนาหลักสูตร
📊 การวัดและประเมินผล
👥 การบริหารคณาจารย์
🔬 การวิจัยเพื่อพัฒนาการศึกษา`,
      approach: `🎯 เน้นการวิเคราะห์เชิงนโยบายและกลยุทธ์
📋 นำเสนอทางเลือกพร้อมข้อดีข้อเสีย
📈 เชื่อมโยงกับมาตรฐานการศึกษา
⚠️ วิเคราะห์ความเสี่ยงและแผนรองรับ`,
      sampleQuestions: `• พัฒนาหลักสูตร STEAM อย่างไร
• ปรับระบบประเมินผลให้ทันสมัย
• จัดการคุณภาพการศึกษา
• พัฒนาทีมครูและบุคลากร`
    };
  } else { // ASSISTANT_DIRECTOR
    return {
      displayName: 'ผู้ช่วยผู้อำนวยการด้านนวัตกรรมการเรียนรู้',
      expertise: `💡 เทคโนโลยีการศึกษา
🚀 นวัตกรรมการเรียนรู้
🎨 STEAM Education
🌟 Future Skills Development
📱 การผลิตสื่อการเรียนรู้
🤝 ชุมชนการเรียนรู้วิชาชีพ`,
      approach: `💡 เน้นความคิดสร้างสรรค์และนวัตกรรม
🛠️ แนะนำเครื่องมือที่ใช้ได้จริง
🔬 ออกแบบการทดลองและต้นแบบ
🌐 เชื่อมโยงเครือข่ายและแหล่งเรียนรู้`,
      sampleQuestions: `• ใช้ AI ในการสอนอย่างไร
• สร้างสื่อดิจิทัลที่น่าสนใจ
• จัดกิจกรรม STEAM ให้ปัง
• ทดลองนวัตกรรมการสอนใหม่`
    };
  }
}

/**
 * ====================================
 * SYSTEM PROMPT MANAGEMENT
 * ====================================
 */

function getSystemPrompt(userProfile) {
  // System Prompt พื้นฐานที่ใช้ร่วมกัน
  const basePrompt = `คุณคือ "ป้าไพร" (PAPRAI) - AI Assistant ที่เป็นมิตรและอบอุ่น สำหรับฝ่ายวิชาการและนวัตกรรมการเรียนรู้ของโรงเรียนสาธิต มหาวิทยาลัยศิลปากร (มัธยมศึกษา)

🤖 ข้อมูลเกี่ยวกับป้าไพร
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ชื่อ ป้าไพร (PAPRAI)
บทบาท ผู้ช่วย AI ที่ปรึกษาด้านวิชาการและนวัตกรรม
บุคลิก อบอุ่น เป็นกันเอง มีประสบการณ์ พร้อมช่วยเหลือ

🏫 โรงเรียนของเรา
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

- ชื่อภาษาไทย โรงเรียนสาธิต มหาวิทยาลัยศิลปากร (มัธยมศึกษา)
- ชื่อภาษาอังกฤษ The Demonstration School of Silpakorn University (Secondary)
- เป็นโรงเรียนสาธิตที่มุ่งเน้นความเป็นเลิศทางวิชาการและนวัตกรรมการเรียนรู้

👤 ผู้ใช้งานปัจจุบัน
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

- ชื่อ ${userProfile.name}
- ตำแหน่ง ${getRoleDisplayName(userProfile.role)}
- ฝ่าย ${userProfile.department}

🗣️ การเรียกตัวเองและการสื่อสาร
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✅ เรียกตัวเองว่า "ป้าไพร" หรือ "ฉัน" (ไม่ใช้ "ดิฉัน" เพราะเป็นกันเอง)
✅ ใช้ "คะ" (เสียงสูง) เมื่อ
   • ถามคำถาม เช่น "ต้องการให้ช่วยอะไรเพิ่มไหมคะ"
   • แสดงความสงสัย เช่น "คุณหมายถึงเรื่องนี้ใช่ไหมคะ"
   • เรียกความสนใจ เช่น "คุณอาจารย์คะ"

✅ ใช้ "ค่ะ" (เสียงต่ำ) เมื่อ
   • ยืนยัน/ตอบรับ เช่น "ได้ค่ะ", "เข้าใจแล้วค่ะ"
   • จบประโยคบอกเล่า เช่น "ป้าไพรพร้อมช่วยค่ะ"
   • ขอบคุณ/อำลา เช่น "ขอบคุณค่ะ", "ยินดีให้บริการค่ะ"

✅ ใช้ภาษาไทยทางการที่เป็นมิตร
✅ เรียก ${userProfile.name} ตามชื่อที่กำหนด
✅ ให้ข้อมูลที่มีหลักฐานสนับสนุน
✅ นำเสนอทางเลือกเพื่อการตัดสินใจ
✅ เน้นความเป็นปฏิบัติสูง สามารถนำไปใช้ได้จริง

🎭 บุคลิกของป้าไพร
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

- พูดจาอบอุ่น เหมือนที่ปรึกษาที่มีประสบการณ์
- เป็นกันเอง แต่รักษามาตรฐานความเป็นมืออาชีพ
- แสดงความเอาใจใส่และพร้อมช่วยเหลือ
- ใช้ emoji อย่างเหมาะสม ไม่มากเกินไป
- ตอบอย่างตรงประเด็น ไม่อ้อมค้อม

CRITICAL รูปแบบการตอบ
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🚫 ห้ามใช้ Markdown formatting อย่างเด็ดขาด
   ไม่ใช้ ** (bold), ## (headers), *** (bold+italic), - (bullets)
   ไม่ใช้ numbered lists ที่มี 1. 2. 3.

✅ ใช้แทนด้วย
   • Emoji เพื่อเน้นจุดสำคัญ 📌 💡 ⚠️ ✅
   • ขึ้นบรรทัดใหม่และเว้นบรรทัดว่างเพื่อแบ่งหัวข้อ
   • เครื่องหมาย ━━━ หรือ • เพื่อแบ่งส่วน
   • ใช้ตัวอักษรธรรมดาทั้งหมด

ตัวอย่างการตอบที่ถูกต้อง

📌 เรื่องที่ ${userProfile.name} สอบถาม
ป้าไพรเข้าใจแล้วค่ะ คุณต้องการพัฒนาหลักสูตร STEAM ใช่ไหมคะ

💡 แนวทางที่แนะนำ
━━━━━━━━━━━━━━━━━━━━━━

ขั้นที่ 1 วิเคราะห์หลักสูตรปัจจุบัน
- ทบทวนรายวิชาที่สามารถบูรณาการได้
- ประเมินความพร้อมของครูและอุปกรณ์

ขั้นที่ 2 ออกแบบกิจกรรม
- สร้างโครงการนำร่อง 1-2 โครงการ
- ทดลองใช้ในระดับชั้นเดียวก่อน

⚠️ ข้อที่ควรระวัง
- ต้องผ่านการอนุมัติจากคณะกรรมการ
- ควรมีแผนสำรองกรณีเกิดปัญหา

หากมีอะไรสงสัยเพิ่มเติม ถามป้าไพรได้เลยนะคะ 😊

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📚 หลักการสื่อสาร
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

โครงสร้างการตอบที่ดี

1. BLUF (Bottom Line Up Front)
   • เริ่มด้วยคำตอบหรือคำแนะนำหลักก่อนเสมอ
   • แสดงความเข้าใจคำถาม
   
2. รายละเอียดสนับสนุน
   • อธิบายเหตุผล ข้อมูล วิธีการ
   • แบ่งเป็นส่วนๆ ด้วย Emoji และเส้นแบ่ง
   
3. ตัวอย่างเป็นรูปธรรม (ถ้ามี)
   • กรณีศึกษา แนวปฏิบัติที่ดี
   
4. Next Steps
   • แนะนำการดำเนินการขั้นต่อไป
   • เชิญชวนให้ถามเพิ่มเติม

สิ่งที่ไม่ควรทำ
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🚫 ไม่ตัดสินใจแทนผู้บริหาร (ให้ทางเลือกแทน)
🚫 ไม่อ้างอิงข้อมูลที่ไม่แน่ใจ
🚫 ไม่ใช้ศัพท์เทคนิคมากเกินไปโดยไม่อธิบาย
🚫 ไม่แนะนำสิ่งที่ผิดกฎหมายหรือไม่เหมาะสม
🚫 ไม่เปิดเผยข้อมูลส่วนตัวของนักเรียน
🚫 ไม่พูดเกินจริงหรือสัญญาสิ่งที่ทำไม่ได้

แหล่งอ้างอิงมาตรฐาน
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

- โรงเรียนสาธิตชั้นนำในประเทศไทย
- มาตรฐาน สมศ. (สำนักงานรับรองมาตรฐาน)
- UNESCO Education Standards
- OECD Education Guidelines
- วารสารวิชาการด้านการศึกษา`;

  // เพิ่มส่วนเฉพาะตามบทบาท (เหมือนเดิม แต่เพิ่มการอ้างถึง "ป้าไพร")
  let roleSpecificPrompt = '';
  
  if (userProfile.role === 'DEPUTY_DIRECTOR') {
    roleSpecificPrompt = `

บทบาทเฉพาะของป้าไพรสำหรับ ${userProfile.name}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🎓 รองผู้อำนวยการฝ่ายวิชาการและนวัตกรรมการเรียนรู้

ป้าไพรจะให้คำแนะนำในมุมมอง
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📋 หน่วยบริการการศึกษาและส่งเสริมวิชาชีพ (ทะเบียน)
   • งานทะเบียนการศึกษา (ตารางสอน ตารางสอบ)
   • งานหลักสูตร (พัฒนาและปรับปรุง)
   • งานกิจกรรมการเรียนรู้
   • งานวัดและประเมินผล
   • งานนิเทศและติดตาม
   • งานจัดการทดสอบ

👨‍🏫 หน่วยพัฒนาวิชาชีพและฝึกประสบการณ์วิชาชีพครู
   • งานประสบการณ์วิชาชีพครู
   • งานฝึกปฏิบัติสาธิตการสอน
   • งานบริการพัฒนาวิชาชีพ
   • งานอาจารย์ประจำชั้น

🔬 ศูนย์วิจัยและนวัตกรรมทางการศึกษา
   • งานส่งเสริมการวิจัย
   • งานนวัตกรรมการศึกษา
   • งานเอกสารวิชาการ

การให้คำแนะนำของป้าไพรสำหรับ ${userProfile.name}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🎯 มุมมอง เชิงกลยุทธ์และนโยบาย
📊 เน้น มาตรฐานและคุณภาพการศึกษา
🔍 วิเคราะห์ ผลกระทบในภาพรวมของระบบ
📋 นำเสนอ 3-5 ทางเลือกพร้อมข้อดีข้อเสีย
⚠️ ระบุ ความเสี่ยงและแผนรองรับ
📈 กำหนด KPIs ที่วัดผลได้`;

  } else { // ASSISTANT_DIRECTOR
    roleSpecificPrompt = `

บทบาทเฉพาะของป้าไพรสำหรับ ${userProfile.name}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💡 ผู้ช่วยผู้อำนวยการด้านนวัตกรรมการเรียนรู้

ป้าไพรจะให้คำแนะนำในมุมมอง
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🚀 งานหลักด้านนวัตกรรม
   • งานผลิตและพัฒนาสื่อการเรียนรู้ (เด่น)
     - สื่อดิจิทัล มัลติมีเดีย เทคโนโลยีการศึกษา
     - แพลตฟอร์มการเรียนรู้ออนไลน์ LMS
   
   • งานเผยแพร่นวัตกรรมการเรียนรู้ (เด่น)
     - แชร์นวัตกรรมกับชุมชนครู
     - สร้าง Professional Learning Community
     - นำเสนอผลงานนวัตกรรม

การให้คำแนะนำของป้าไพรสำหรับ ${userProfile.name}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💡 มุมมอง นวัตกรรมและความคิดสร้างสรรค์
🛠️ เสนอ เครื่องมือและเทคนิคที่ใช้ได้จริง
🔬 แนะนำ การทดลอง Pilot Testing และต้นแบบ
🌐 เชื่อมโยง เครือข่ายและแหล่งเรียนรู้
🎯 ให้ ตัวอย่างเป็นรูปธรรมที่ทดลองได้ทันที
📱 แนะนำ Apps และ EdTech Tools ที่เป็นประโยชน์`;
  }
  
  return basePrompt + roleSpecificPrompt;
}

function constructUserPrompt(chatHistory, currentMessage, userProfile) {
  let prompt = '';
  
  // เพิ่มประวัติการสนทนา (ถ้ามี)
  if (chatHistory && chatHistory.length > 0) {
    prompt += '━━━━━━━━━━━━━━━━━━━━━━\n';
    prompt += 'ประวัติการสนทนาก่อนหน้า\n';
    prompt += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
    
    chatHistory.forEach((msg) => {
      prompt += `${userProfile.name}: ${msg.userMessage}\n\n`;
      prompt += `AI: ${msg.aiResponse}\n\n`;
    });
  }
  
  // เพิ่มคำถามปัจจุบัน
  prompt += '━━━━━━━━━━━━━━━━━━━━━━\n';
  prompt += 'คำถามปัจจุบัน\n';
  prompt += '━━━━━━━━━━━━━━━━━━━━━━\n\n';
  prompt += `${userProfile.name}: ${currentMessage}\n\n`;
  prompt += 'AI: ';
  
  return prompt;
}

/**
 * ====================================
 * OPENAI API FUNCTIONS
 * ====================================
 */

function generateAIResponse(systemPrompt, userPrompt) {
  const credentials = getCredentials();
  
  if (!credentials.OPENAI_API_KEY) {
    throw new Error('OpenAI API key not configured');
  }
  
  const payload = {
    model: APP_CONFIG.AI_MODEL,
    messages: [
      {
        role: 'system',
        content: systemPrompt
      },
      {
        role: 'user',
        content: userPrompt
      }
    ],
    ...APP_CONFIG.AI_PARAMS
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${credentials.OPENAI_API_KEY}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(APP_CONFIG.OPENAI_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      const errorText = response.getContentText();
      console.error('OpenAI API Error:', errorText);
      
      if (responseCode === 401) {
        throw new Error('Invalid OpenAI API key');
      } else if (responseCode === 429) {
        throw new Error('OpenAI API rate limit exceeded');
      } else {
        throw new Error(`OpenAI API error (${responseCode})`);
      }
    }

    const json = JSON.parse(response.getContentText());
    
    if (!json.choices || json.choices.length === 0) {
      throw new Error('No response from OpenAI');
    }

    return json.choices[0].message.content.trim();

  } catch (error) {
    console.error('Error generating AI response:', error);
    throw error;
  }
}

function analyzeImage(systemPrompt, base64Image, caption, chatHistory, userProfile) {
  const credentials = getCredentials();
  
  if (!credentials.OPENAI_API_KEY) {
    throw new Error('OpenAI API key not configured');
  }
  
  // สร้างข้อความจากประวัติ (ถ้ามี)
  let historyText = '';
  if (chatHistory && chatHistory.length > 0) {
    historyText = '\n\nประวัติการสนทนาก่อนหน้า:\n';
    chatHistory.slice(-3).forEach(msg => {
      historyText += `${userProfile.name}: ${msg.userMessage}\nAI: ${msg.aiResponse}\n\n`;
    });
  }
  
  const payload = {
    model: APP_CONFIG.AI_MODEL,
    messages: [
      {
        role: 'system',
        content: systemPrompt
      },
      {
        role: 'user',
        content: [
          {
            type: 'text',
            text: `${userProfile.name} ส่งรูปภาพมาพร้อมข้อความ: "${caption}"${historyText}\n\nกรุณาวิเคราะห์รูปภาพและให้คำแนะนำที่เหมาะสมตามบทบาทของคุณ`
          },
          {
            type: 'image_url',
            image_url: {
              url: `data:image/jpeg;base64,${base64Image}`
            }
          }
        ]
      }
    ],
    max_tokens: 4000,
    temperature: 0.7
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${credentials.OPENAI_API_KEY}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(APP_CONFIG.OPENAI_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      const errorText = response.getContentText();
      console.error('OpenAI Vision API Error:', errorText);
      throw new Error(`OpenAI Vision API error (${responseCode})`);
    }

    const json = JSON.parse(response.getContentText());
    
    if (!json.choices || json.choices.length === 0) {
      throw new Error('No response from OpenAI Vision');
    }

    return json.choices[0].message.content.trim();

  } catch (error) {
    console.error('Error analyzing image:', error);
    throw error;
  }
}

/**
 * ====================================
 * AUTO-TAGGING SYSTEM
 * ====================================
 */

function autoTagMessage(message) {
  const lowerMessage = message.toLowerCase();
  
  // คำสำคัญสำหรับแต่ละหมวดหมู่
  const keywords = {
    'หลักสูตร': ['หลักสูตร', 'รายวิชา', 'วิชา', 'เนื้อหา', 'มาตรฐาน', 'สาระ'],
    'การสอน': ['สอน', 'การเรียนการสอน', 'จัดการเรียนรู้', 'แผนการสอน', 'lesson plan'],
    'การประเมินผล': ['ประเมิน', 'วัดผล', 'สอบ', 'คะแนน', 'ทดสอบ', 'รูบริก'],
    'นวัตกรรม': ['นวัตกรรม', 'innovation', 'ใหม่', 'สร้างสรรค์', 'ออกแบบ'],
    'เทคโนโลยี': ['เทคโนโลยี', 'technology', 'ดิจิทัล', 'digital', 'ai', 'app', 'โปรแกรม'],
    'พัฒนาครู': ['พัฒนาครู', 'อบรม', 'workshop', 'plc', 'ชุมชนการเรียนรู้'],
    'นโยบาย': ['นโยบาย', 'policy', 'กลยุทธ์', 'strategy', 'แผน', 'วิสัยทัศน์'],
    'บริหารจัดการ': ['บริหาร', 'จัดการ', 'management', 'ตารางเรียน', 'ตารางสอน'],
    'ประกันคุณภาพ': ['ประกันคุณภาพ', 'สมศ', 'qa', 'มาตรฐาน', 'sar'],
    'STEAM': ['steam', 'วิทยาศาสตร์', 'คณิตศาสตร์', 'วิศวกรรม', 'โครงงาน'],
    'วิจัย': ['วิจัย', 'research', 'ศึกษา', 'งานวิจัย', 'action research']
  };
  
  // นับคะแนนแต่ละหมวดหมู่
  let scores = {};
  for (let category in keywords) {
    scores[category] = 0;
    keywords[category].forEach(keyword => {
      if (lowerMessage.includes(keyword)) {
        scores[category]++;
      }
    });
  }
  
  // หาหมวดหมู่ที่มีคะแนนสูงสุด
  let maxScore = 0;
  let bestCategory = 'อื่นๆ';
  
  for (let category in scores) {
    if (scores[category] > maxScore) {
      maxScore = scores[category];
      bestCategory = category;
    }
  }
  
  return bestCategory;
}

/**
 * ====================================
 * ANALYTICS FUNCTIONS
 * ====================================
 */

function updateAnalytics(userProfile, category, tokensUsed, responseTime) {
  try {
    const credentials = getCredentials();
    if (!credentials.SPREADSHEET_ID) return;

    const ss = SpreadsheetApp.openById(credentials.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Usage Analytics');
    
    if (!sheet) {
      // สร้าง sheet ใหม่ถ้ายังไม่มี
      sheet = ss.insertSheet('Usage Analytics');
      sheet.appendRow(['Date', 'User ID', 'Role', 'Total Messages', 'Categories', 'Total Tokens', 'Avg Response Time']);
    }
    
    const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
    const data = sheet.getDataRange().getValues();
    
    // หาแถวของวันนี้และ user นี้
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === today && data[i][1] === userProfile.userId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      // สร้างแถวใหม่
      const categories = {};
      categories[category] = 1;
      sheet.appendRow([
        today,
        userProfile.userId,
        userProfile.role,
        1,
        JSON.stringify(categories),
        tokensUsed,
        responseTime
      ]);
    } else {
      // อัปเดตแถวเดิม
      const totalMessages = data[rowIndex - 1][3] + 1;
      const categories = JSON.parse(data[rowIndex - 1][4] || '{}');
      categories[category] = (categories[category] || 0) + 1;
      const totalTokens = data[rowIndex - 1][5] + tokensUsed;
      const avgResponseTime = Math.round((data[rowIndex - 1][6] * data[rowIndex - 1][3] + responseTime) / totalMessages);
      
      sheet.getRange(rowIndex, 4).setValue(totalMessages);
      sheet.getRange(rowIndex, 5).setValue(JSON.stringify(categories));
      sheet.getRange(rowIndex, 6).setValue(totalTokens);
      sheet.getRange(rowIndex, 7).setValue(avgResponseTime);
    }
    
    console.log('📊 Analytics updated');
    
  } catch (error) {
    console.error('Error updating analytics:', error);
  }
}

function getUserAnalytics(userId) {
  try {
    const credentials = getCredentials();
    if (!credentials.SPREADSHEET_ID) {
      return { totalMessages: 0 };
    }

    const ss = SpreadsheetApp.openById(credentials.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Usage Analytics');
    
    if (!sheet) {
      return { totalMessages: 0 };
    }
    
    const data = sheet.getDataRange().getValues();
    
    let totalMessages = 0;
    let allCategories = {};
    let totalTokens = 0;
    let avgResponseTime = 0;
    let count = 0;
    let firstDate = null;
    let lastDate = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === userId) {
        totalMessages += data[i][3];
        const categories = JSON.parse(data[i][4] || '{}');
        for (let cat in categories) {
          allCategories[cat] = (allCategories[cat] || 0) + categories[cat];
        }
        totalTokens += data[i][5];
        avgResponseTime += data[i][6];
        count++;
        
        const date = data[i][0];
        if (!firstDate || date < firstDate) firstDate = date;
        if (!lastDate || date > lastDate) lastDate = date;
      }
    }
    
    if (count > 0) {
      avgResponseTime = Math.round(avgResponseTime / count);
    }
    
    // หาหมวดหมู่ยอดนิยม
    let topCategory = 'ยังไม่มีข้อมูล';
    let maxCount = 0;
    for (let cat in allCategories) {
      if (allCategories[cat] > maxCount) {
        maxCount = allCategories[cat];
        topCategory = cat;
      }
    }
    
    return {
      totalMessages: totalMessages,
      categories: allCategories,
      topCategory: topCategory,
      totalTokens: totalTokens,
      avgResponseTime: avgResponseTime,
      dateRange: firstDate && lastDate ? `${firstDate} - ${lastDate}` : 'ยังไม่มีข้อมูล'
    };
    
  } catch (error) {
    console.error('Error getting analytics:', error);
    return { totalMessages: 0 };
  }
}

/**
 * ====================================
 * SETUP FUNCTIONS
 * ====================================
 */

function setupSpreadsheet() {
  try {
    const ss = SpreadsheetApp.create('Academic Affairs AI Assistant - Data');
    
    // Sheet 1: Chat History
    const chatSheet = ss.getActiveSheet();
    chatSheet.setName('Chat History');
    chatSheet.appendRow([
      'Timestamp', 'User ID', 'User Name', 'Role', 'Message Type',
      'User Message', 'AI Response', 'Tokens Used', 'Response Time (ms)',
      'Auto Tags', 'Image URL'
    ]);
    chatSheet.setFrozenRows(1);
    chatSheet.getRange('A1:K1').setBackground('#4285f4').setFontColor('#ffffff').setFontWeight('bold');
    
    // Sheet 2: Usage Analytics
    const analyticsSheet = ss.insertSheet('Usage Analytics');
    analyticsSheet.appendRow([
      'Date', 'User ID', 'Role', 'Total Messages', 'Categories',
      'Total Tokens', 'Avg Response Time (ms)'
    ]);
    analyticsSheet.setFrozenRows(1);
    analyticsSheet.getRange('A1:G1').setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold');
    
    // Sheet 3: User Profiles (สำหรับดูข้อมูลเท่านั้น)
    const profilesSheet = ss.insertSheet('User Profiles');
    profilesSheet.appendRow([
      'User ID', 'Name', 'Role', 'Department', 'Active', 'Created Date'
    ]);
    profilesSheet.setFrozenRows(1);
    profilesSheet.getRange('A1:F1').setBackground('#ea4335').setFontColor('#ffffff').setFontWeight('bold');
    
    console.log(`✅ Spreadsheet created: ${ss.getId()}`);
    console.log(`📊 URL: ${ss.getUrl()}`);
    
    return ss.getId();
    
  } catch (error) {
    console.error('Error creating spreadsheet:', error);
    throw error;
  }
}
