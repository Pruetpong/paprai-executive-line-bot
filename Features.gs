// Features.gs - Special Features and Automation for Academic Affairs AI Assistant
// ฟีเจอร์พิเศษและระบบอัตโนมัติ

/**
 * ====================================
 * WEEKLY REPORT SYSTEM
 * ====================================
 */

/**
 * ฟังก์ชันส่งรายงานประจำสัปดาห์
 * 
 * วิธีตั้งค่า Trigger:
 * 1. ไปที่ Extensions > Apps Script > Triggers (นาฬิกา)
 * 2. Add Trigger
 * 3. Choose function: sendWeeklyReports
 * 4. Select event source: Time-driven
 * 5. Select type of time based trigger: Week timer
 * 6. Select day of week: Monday
 * 7. Select time of day: 8am to 9am
 * 8. Save
 */
function sendWeeklyReports() {
  console.log('📊 Starting weekly reports generation...');
  
  try {
    const credentials = getCredentials();
    const userProfiles = credentials.USER_PROFILES || {};
    
    // ส่งรายงานให้ผู้ใช้แต่ละคน
    for (let userId in userProfiles) {
      const profile = userProfiles[userId];
      if (profile.active) {
        sendWeeklyReportToUser(userId, profile);
      }
    }
    
    console.log('✅ Weekly reports sent successfully');
    
  } catch (error) {
    console.error('❌ Error sending weekly reports:', error);
  }
}

function sendWeeklyReportToUser(userId, profile) {
  try {
    console.log(`📧 Generating report for ${profile.name}...`);
    
    // คำนวณช่วงเวลา (สัปดาห์ที่แล้ว)
    const today = new Date();
    const lastMonday = new Date(today);
    lastMonday.setDate(today.getDate() - 7);
    const lastSunday = new Date(today);
    lastSunday.setDate(today.getDate() - 1);
    
    // ดึงข้อมูลการใช้งานสัปดาห์ที่แล้ว
    const weekData = getWeeklyData(userId, lastMonday, lastSunday);
    
    if (weekData.totalMessages === 0) {
      console.log(`⚠️ No activity for ${profile.name} last week`);
      return;
    }
    
    // สร้างรายงาน
    const reportText = generateWeeklyReportText(profile, weekData, lastMonday, lastSunday);
    
    // บันทึกรายงาน
    saveWeeklyReport(userId, reportText, lastMonday, lastSunday);
    
    // ส่งรายงานผ่าน LINE
    const messages = [{
      type: 'text',
      text: reportText
    }];
    
    pushMessage(userId, messages);
    
    console.log(`✅ Report sent to ${profile.name}`);
    
  } catch (error) {
    console.error(`❌ Error sending report to ${profile.name}:`, error);
  }
}

function getWeeklyData(userId, startDate, endDate) {
  try {
    const credentials = getCredentials();
    if (!credentials.SPREADSHEET_ID) {
      return { totalMessages: 0 };
    }

    const ss = SpreadsheetApp.openById(credentials.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Chat History');
    
    if (!sheet) {
      return { totalMessages: 0 };
    }
    
    const data = sheet.getDataRange().getValues();
    
    let totalMessages = 0;
    let categories = {};
    let totalResponseTime = 0;
    let imageCount = 0;
    let recentTopics = [];
    
    for (let i = 1; i < data.length; i++) {
      const rowUserId = data[i][1];
      const timestamp = new Date(data[i][0]);
      
      if (rowUserId === userId && timestamp >= startDate && timestamp <= endDate) {
        totalMessages++;
        
        const category = data[i][9]; // Column J = Auto Tags
        categories[category] = (categories[category] || 0) + 1;
        
        const responseTime = data[i][8]; // Column I = Response Time
        totalResponseTime += responseTime || 0;
        
        const messageType = data[i][4]; // Column E = Message Type
        if (messageType === 'image') {
          imageCount++;
        }
        
        const userMessage = data[i][5]; // Column F = User Message
        if (recentTopics.length < 5) {
          recentTopics.push(userMessage.substring(0, 50) + (userMessage.length > 50 ? '...' : ''));
        }
      }
    }
    
    const avgResponseTime = totalMessages > 0 ? Math.round(totalResponseTime / totalMessages) : 0;
    
    // หาหมวดหมู่ยอดนิยม 3 อันดับแรก
    const topCategories = Object.entries(categories)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([cat, count]) => `${cat} (${count} ครั้ง)`);
    
    return {
      totalMessages: totalMessages,
      categories: categories,
      topCategories: topCategories,
      avgResponseTime: avgResponseTime,
      imageCount: imageCount,
      recentTopics: recentTopics
    };
    
  } catch (error) {
    console.error('Error getting weekly data:', error);
    return { totalMessages: 0 };
  }
}

function generateWeeklyReportText(profile, weekData, startDate, endDate) {
  const formatDate = (date) => {
    return Utilities.formatDate(date, 'Asia/Bangkok', 'd MMM yyyy', 'th-TH');
  };
  
  let reportText = `📊 รายงานสรุปการใช้งานประจำสัปดาห์

สวัสดีค่ะ ${profile.name} 🙏
ป้าไพรส่งรายงานมาให้แล้วค่ะ

นี่คือสรุปการใช้งานช่วง
${formatDate(startDate)} - ${formatDate(endDate)}

━━━━━━━━━━━━━━━━━━━━━━

📈 สถิติการใช้งาน

💬 จำนวนข้อความทั้งหมด
   ${weekData.totalMessages} ข้อความ

📷 รูปภาพที่วิเคราะห์
   ${weekData.imageCount} รูป

⏱️ เวลาตอบกลับเฉลี่ย
   ${weekData.avgResponseTime} มิลลิวินาที

━━━━━━━━━━━━━━━━━━━━━━

🏷️ หมวดหมู่ที่ปรึกษาบ่อยสุด

`;

  weekData.topCategories.forEach((cat, i) => {
    reportText += `${i + 1}. ${cat}\n`;
  });
  
  if (weekData.recentTopics.length > 0) {
    reportText += `\n━━━━━━━━━━━━━━━━━━━━━━\n\n`;
    reportText += `💡 หัวข้อที่ปรึกษาล่าสุด\n\n`;
    
    weekData.recentTopics.forEach((topic, i) => {
      reportText += `• ${topic}\n`;
    });
  }
  
  reportText += `\n━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  reportText += `📝 ข้อสังเกตของป้าไพร\n\n`;
  
  // ให้คำแนะนำตามข้อมูล
  if (weekData.totalMessages < 5) {
    reportText += `ป้าไพรแนะนำให้ใช้งานบ่อยๆ นะคะ\nจะได้ช่วยงานด้านวิชาการ\nของคุณได้มากขึ้นค่ะ\n`;
  } else if (weekData.totalMessages > 20) {
    reportText += `คุณใช้งานบ่อยมากเลยค่ะ 👍\nป้าไพรหวังว่าจะช่วยให้งาน\nของคุณสะดวกขึ้นนะคะ\n`;
  } else {
    reportText += `การใช้งานดีมากค่ะ ✨\nมีอะไรสงสัยเพิ่มเติม\nถามป้าไพรได้เลยนะคะ\n`;
  }
  
  reportText += `\n━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  reportText += `🔍 ดูรายละเอียดเพิ่มเติม\n`;
  reportText += `พิมพ์ /analytics\n\n`;
  reportText += `ขอบคุณที่ใช้บริการป้าไพรค่ะ 🙏`;
  
  return reportText;
}

function saveWeeklyReport(userId, reportContent, startDate, endDate) {
  try {
    const credentials = getCredentials();
    if (!credentials.SPREADSHEET_ID) return;

    const ss = SpreadsheetApp.openById(credentials.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Weekly Reports');
    
    if (!sheet) return;
    
    sheet.appendRow([
      Utilities.formatDate(startDate, 'Asia/Bangkok', 'yyyy-MM-dd'),
      Utilities.formatDate(endDate, 'Asia/Bangkok', 'yyyy-MM-dd'),
      userId,
      reportContent,
      new Date(),
      'SENT'
    ]);
    
    console.log('📝 Weekly report saved to spreadsheet');
    
  } catch (error) {
    console.error('Error saving weekly report:', error);
  }
}

/**
 * ====================================
 * SMART ALERTS SYSTEM
 * ====================================
 */

/**
 * ฟังก์ชันส่งการแจ้งเตือนเช้า (Morning Brief)
 * 
 * วิธีตั้งค่า Trigger:
 * 1. Add Trigger
 * 2. Choose function: sendMorningBrief
 * 3. Select event source: Time-driven
 * 4. Select type of time based trigger: Day timer
 * 5. Select time of day: 7am to 8am
 * 6. Save
 */
function sendMorningBrief() {
  console.log('🌅 Sending morning brief...');
  
  try {
    const credentials = getCredentials();
    const userProfiles = credentials.USER_PROFILES || {};
    
    for (let userId in userProfiles) {
      const profile = userProfiles[userId];
      if (profile.active) {
        sendMorningBriefToUser(userId, profile);
      }
    }
    
    console.log('✅ Morning briefs sent');
    
  } catch (error) {
    console.error('❌ Error sending morning briefs:', error);
  }
}

function sendMorningBriefToUser(userId, profile) {
  try {
    // ตรวจสอบว่ามีข้อความเมื่อวานหรือไม่
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    yesterday.setHours(0, 0, 0, 0);
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const yesterdayData = getWeeklyData(userId, yesterday, today);
    
    if (yesterdayData.totalMessages === 0) {
      // ไม่ส่งถ้าไม่มีกิจกรรมเมื่อวาน
      return;
    }
    
    const briefText = `🌅 สวัสดีตอนเช้าค่ะ ${profile.name}

💡 สรุปย่อเมื่อวาน
   คุณปรึกษาเรื่อง ${yesterdayData.topCategories[0] || 'ต่างๆ'}
   รวม ${yesterdayData.totalMessages} ข้อความ

✨ วันนี้ป้าไพรพร้อมให้คำปรึกษาแล้วค่ะ
   มีอะไรสงสัย พิมพ์มาได้เลยนะคะ 😊

━━━━━━━━━━━━━━━━━━━━━━

⚡ เคล็ดลับวันนี้
   ลองใช้คำสั่ง /analytics
   เพื่อดูสถิติการใช้งานโดยรวม

ขอให้เป็นวันที่ดีนะคะ 🙏`;
    
    const messages = [{ type: 'text', text: briefText }];
    pushMessage(userId, messages);
    
    console.log(`✅ Morning brief sent to ${profile.name}`);
    
  } catch (error) {
    console.error(`❌ Error sending morning brief to ${profile.name}:`, error);
  }
}

/**
 * ฟังก์ชันส่งสรุปตอนเย็น (Evening Summary)
 * 
 * วิธีตั้งค่า Trigger:
 * 1. Add Trigger
 * 2. Choose function: sendEveningSummary
 * 3. Select event source: Time-driven
 * 4. Select type of time based trigger: Day timer
 * 5. Select time of day: 5pm to 6pm
 * 6. Save
 */
function sendEveningSummary() {
  console.log('🌆 Sending evening summary...');
  
  try {
    const credentials = getCredentials();
    const userProfiles = credentials.USER_PROFILES || {};
    
    for (let userId in userProfiles) {
      const profile = userProfiles[userId];
      if (profile.active) {
        sendEveningSummaryToUser(userId, profile);
      }
    }
    
    console.log('✅ Evening summaries sent');
    
  } catch (error) {
    console.error('❌ Error sending evening summaries:', error);
  }
}

function sendEveningSummaryToUser(userId, profile) {
  try {
    // ดึงข้อมูลวันนี้
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const now = new Date();
    
    const todayData = getWeeklyData(userId, today, now);
    
    if (todayData.totalMessages === 0) {
      // ไม่ส่งถ้าไม่มีกิจกรรมวันนี้
      return;
    }
    
    const summaryText = `🌆 สรุปการทำงานวันนี้

สวัสดีตอนเย็นค่ะ ${profile.name} 🙏

📊 สรุปการใช้งานวันนี้
   • ข้อความทั้งหมด ${todayData.totalMessages} ข้อความ
   • หมวดหมู่หลัก ${todayData.topCategories.join(', ')}
   ${todayData.imageCount > 0 ? `• วิเคราะห์รูปภาพ ${todayData.imageCount} รูป` : ''}

━━━━━━━━━━━━━━━━━━━━━━

💡 คำถามที่ปรึกษาวันนี้

${todayData.recentTopics.slice(0, 3).map((topic, i) => `${i + 1}. ${topic}`).join('\n')}

━━━━━━━━━━━━━━━━━━━━━━

✨ ยังมีอะไรค้างคาอยู่ไหมคะ
   ป้าไพรพร้อมช่วยตลอดนะคะ

พักผ่อนให้เพียงพอนะคะ 😊
ราตรีสวัสดิ์ค่ะ 🌙`;
    
    const messages = [{ type: 'text', text: summaryText }];
    pushMessage(userId, messages);
    
    console.log(`✅ Evening summary sent to ${profile.name}`);
    
  } catch (error) {
    console.error(`❌ Error sending evening summary to ${profile.name}:`, error);
  }
}

/**
 * ====================================
 * ADVANCED ANALYTICS
 * ====================================
 */

function generateMonthlyInsights(userId) {
  try {
    const credentials = getCredentials();
    if (!credentials.SPREADSHEET_ID) {
      return "ยังไม่ได้ตั้งค่า Spreadsheet ค่ะ";
    }

    // ดึงข้อมูล 30 วันที่แล้ว
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(endDate.getDate() - 30);
    
    const monthData = getWeeklyData(userId, startDate, endDate);
    
    if (monthData.totalMessages === 0) {
      return "ยังไม่มีข้อมูลการใช้งานในเดือนนี้ค่ะ";
    }
    
    // วิเคราะห์เทรนด์
    const insights = [];
    
    if (monthData.totalMessages > 100) {
      insights.push("คุณใช้งานบ่อยมาก แสดงว่า AI Assistant มีประโยชน์ต่อคุณ 🌟");
    } else if (monthData.totalMessages > 50) {
      insights.push("การใช้งานอยู่ในระดับที่ดี ต่อเนื่องและสม่ำเสมอ ✨");
    } else {
      insights.push("แนะนำให้ลองใช้งานบ่อยขึ้น เพื่อช่วยเพิ่มประสิทธิภาพการทำงาน 💪");
    }
    
    // หมวดหมู่ที่โดดเด่น
    const topCat = monthData.topCategories[0] || 'ไม่มี';
    insights.push(`คุณสนใจเรื่อง ${topCat.split(' ')[0]} มากที่สุด`);
    
    // ความเร็วในการตอบ
    if (monthData.avgResponseTime < 3000) {
      insights.push("ระบบตอบกลับเร็วมาก ประสิทธิภาพดีเยี่ยม 🚀");
    } else if (monthData.avgResponseTime < 5000) {
      insights.push("เวลาตอบกลับอยู่ในเกณฑ์ดี ⚡");
    }
    
    return insights.join('\n\n');
    
  } catch (error) {
    console.error('Error generating insights:', error);
    return "ไม่สามารถสร้างรายงานได้ค่ะ 🙏";
  }
}

/**
 * ====================================
 * UTILITY FUNCTIONS
 * ====================================
 */

/**
 * ฟังก์ชันสำหรับทดสอบการส่งข้อความ
 */
function testPushMessage() {
  const credentials = getCredentials();
  const userProfiles = credentials.USER_PROFILES || {};
  
  const firstUserId = Object.keys(userProfiles)[0];
  const profile = userProfiles[firstUserId];
  
  if (!firstUserId) {
    console.error('❌ No users configured');
    return;
  }
  
  const testMessage = `🧪 ทดสอบการส่งข้อความ

สวัสดีค่ะ ${profile.name}

นี่เป็นข้อความทดสอบจากป้าไพร
ถ้าคุณเห็นข้อความนี้ แสดงว่า
ระบบการส่งข้อความทำงานปกติค่ะ ✅

มีอะไรให้ป้าไพรช่วยไหมคะ 🙏`;
  
  const messages = [{ type: 'text', text: testMessage }];
  
  try {
    pushMessage(firstUserId, messages);
    console.log('✅ Test message sent successfully');
  } catch (error) {
    console.error('❌ Failed to send test message:', error);
  }
}

/**
 * ฟังก์ชันลบข้อมูลเก่า (เรียกทุก 3 เดือน)
 * 
 * วิธีตั้งค่า Trigger:
 * 1. Add Trigger
 * 2. Choose function: cleanOldData
 * 3. Select event source: Time-driven
 * 4. Select type of time based trigger: Month timer
 * 5. Select day of month: 1
 * 6. Save
 */
function cleanOldData() {
  console.log('🧹 Starting data cleanup...');
  
  try {
    const credentials = getCredentials();
    if (!credentials.SPREADSHEET_ID) {
      console.log('⚠️ No spreadsheet configured');
      return;
    }

    const ss = SpreadsheetApp.openById(credentials.SPREADSHEET_ID);
    const chatSheet = ss.getSheetByName('Chat History');
    const analyticsSheet = ss.getSheetByName('Usage Analytics');
    
    // ลบข้อมูลที่เก่ากว่า 6 เดือน
    const sixMonthsAgo = new Date();
    sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
    
    if (chatSheet) {
      const data = chatSheet.getDataRange().getValues();
      const rowsToDelete = [];
      
      for (let i = data.length - 1; i >= 1; i--) {
        const timestamp = new Date(data[i][0]);
        if (timestamp < sixMonthsAgo) {
          rowsToDelete.push(i + 1);
        }
      }
      
      // ลบทีละแถวจากล่างขึ้นบน
      rowsToDelete.forEach(rowNum => {
        chatSheet.deleteRow(rowNum);
      });
      
      console.log(`🗑️ Deleted ${rowsToDelete.length} old chat history entries`);
    }
    
    if (analyticsSheet) {
      const data = analyticsSheet.getDataRange().getValues();
      const rowsToDelete = [];
      
      for (let i = data.length - 1; i >= 1; i--) {
        const date = new Date(data[i][0]);
        if (date < sixMonthsAgo) {
          rowsToDelete.push(i + 1);
        }
      }
      
      rowsToDelete.forEach(rowNum => {
        analyticsSheet.deleteRow(rowNum);
      });
      
      console.log(`🗑️ Deleted ${rowsToDelete.length} old analytics entries`);
    }
    
    console.log('✅ Data cleanup completed');
    
  } catch (error) {
    console.error('❌ Error cleaning old data:', error);
  }
}

/**
 * ====================================
 * MANUAL TRIGGERS (สำหรับทดสอบ)
 * ====================================
 */

/**
 * ทดสอบส่งรายงานสัปดาห์
 */
function testWeeklyReport() {
  console.log('🧪 Testing weekly report...');
  sendWeeklyReports();
}

/**
 * ทดสอบส่ง Morning Brief
 */
function testMorningBrief() {
  console.log('🧪 Testing morning brief...');
  sendMorningBrief();
}

/**
 * ทดสอบส่ง Evening Summary
 */
function testEveningSummary() {
  console.log('🧪 Testing evening summary...');
  sendEveningSummary();
}

/**
 * ดูสถิติการใช้งานทั้งหมด
 */
function viewAllStatistics() {
  const credentials = getCredentials();
  const userProfiles = credentials.USER_PROFILES || {};
  
  console.log('📊 === STATISTICS FOR ALL USERS ===');
  
  for (let userId in userProfiles) {
    const profile = userProfiles[userId];
    const analytics = getUserAnalytics(userId);
    
    console.log(`\n👤 ${profile.name} (${profile.role})`);
    console.log(`   Total Messages: ${analytics.totalMessages}`);
    console.log(`   Top Category: ${analytics.topCategory}`);
    console.log(`   Avg Response Time: ${analytics.avgResponseTime}ms`);
    console.log(`   Date Range: ${analytics.dateRange}`);
  }
  
  console.log('\n✅ Statistics displayed');
}
