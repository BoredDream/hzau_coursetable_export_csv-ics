(function () {
  console.log("=== 正在生成课表 HTML 预览及多格式导出文件 ===");

  // 1. 核心配置：开学日期与精准时间映射
  const TERM_START = new Date('2026-03-02'); 
  const START_TIMES = {
    1: "08:00", 2: "08:55", 3: "10:00", 4: "10:55",
    5: "14:30", 6: "15:25", 7: "16:30", 8: "17:25",
    9: "19:00", 10: "19:50", 11: "20:40", 12: "21:30"
  };
  const END_TIMES = {
    1: "08:45", 2: "09:40", 3: "10:45", 4: "11:40",
    5: "15:15", 6: "16:10", 7: "17:15", 8: "18:10",
    9: "19:45", 10: "20:35", 11: "21:25", 12: "22:15"
  };

  const table = document.querySelector('.schedule-table table');
  if (!table) {
    console.error("❌ 未找到课表表格，请检查是否在正确的 iframe 中运行。");
    return;
  }

  const finalRecords = [];
  const rows = [...table.querySelectorAll('tr')].slice(1);

  // 2. 深度解析逻辑
  rows.forEach((tr) => {
    const cells = tr.querySelectorAll('td');
    cells.forEach((td, colIndex) => {
      if (colIndex === 0) return; // 跳过时间轴列
      const weekday = colIndex; // 1=周一

      // 提取单元格内所有的课程块
      const items = td.querySelectorAll('.course-item');
      items.forEach((item) => {
        const name = (item.querySelector('.name span:last-child') || item.querySelector('.name')).innerText.trim();
        const timeText = item.querySelector('.el-icon-time + span')?.innerText.trim() || '';
        const location = item.querySelector('.el-icon-location-outline + span')?.innerText.trim() || '';

        const secMatch = timeText.match(/[（\(](\d+)-?(\d*)节[）\)]/);
        if (!secMatch) return;
        const startSec = parseInt(secMatch[1]);
        const endSec = secMatch[2] ? parseInt(secMatch[2]) : startSec;

        const weekPart = timeText.replace(/^[（\(].*?节[）\)]\s*/, '');
        const segments = weekPart.split(/[,，]/);

        segments.forEach(seg => {
          let parity = 0; // 0:全周, 1:单周, 2:双周
          if (seg.includes('单')) parity = 1;
          if (seg.includes('双')) parity = 2;

          const rangeMatch = seg.match(/(\d+)-?(\d*)/);
          if (!rangeMatch) return;
          const startW = parseInt(rangeMatch[1]);
          const endW = rangeMatch[2] ? parseInt(rangeMatch[2]) : startW;

          for (let w = startW; w <= endW; w++) {
            if (parity === 1 && w % 2 === 0) continue;
            if (parity === 2 && w % 2 !== 0) continue;

            let d = new Date(TERM_START);
            d.setDate(TERM_START.getDate() + (w - 1) * 7 + (weekday - 1));
            const dateStr = d.toISOString().split('T')[0].replace(/-/g, '');

            finalRecords.push({
              Subject: name,
              StartDate: d,
              DateRaw: d.toISOString().split('T')[0],
              StartT: START_TIMES[startSec],
              EndT: END_TIMES[endSec],
              Location: location,
              Description: `第${w}周`
            });
          }
        });
      });
    });
  });

  // 3. 构建 HTML 预览与导出界面
  const container = document.createElement('div');
  container.id = "course-exporter-modal";
  container.style = "position:fixed;top:5%;left:5%;width:90%;height:90%;background:white;z-index:99999;overflow:auto;border:3px solid #007131;padding:25px;box-shadow:0 0 20px rgba(0,0,0,0.3);font-family:sans-serif;border-radius:8px;";
  
  let html = `<h2 style="color:#007131">课表解析成功 (共 ${finalRecords.length} 个日程节点)</h2>`;
  html += `<div style="margin-bottom:15px;">
             <button id="dlCSV" style="padding:10px 15px;background:#4CAF50;color:white;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">导出为 CSV (Google日历)</button>
             <button id="dlICS" style="margin-left:10px;padding:10px 15px;background:#2196F3;color:white;border:none;border-radius:4px;cursor:pointer;font-weight:bold;">导出为 ICS (通用格式)</button>
             <button id="closeM" style="margin-left:10px;padding:10px 15px;background:#9e9e9e;color:white;border:none;border-radius:4px;cursor:pointer;">关闭预览</button>
           </div>`;
  html += `<table border="1" style="width:100%;border-collapse:collapse;font-size:13px;">
             <tr style="background:#f2f2f2"><th>课程名称</th><th>日期</th><th>时间段</th><th>地点</th><th>备注</th></tr>`;
  
  finalRecords.forEach(r => {
    html += `<tr><td style="padding:5px">${r.Subject}</td><td>${r.DateRaw}</td><td>${r.StartT}-${r.EndT}</td><td>${r.Location}</td><td>${r.Description}</td></tr>`;
  });
  html += `</table>`;
  
  container.innerHTML = html;
  document.body.appendChild(container);

  // 4. 导出逻辑函数
  const formatDateICS = (date, timeStr) => {
    const [hh, mm] = timeStr.split(':');
    const d = new Date(date);
    d.setHours(hh, mm, 0);
    return d.toISOString().replace(/[-:]/g, '').split('.')[0] + 'Z';
  };

  // 导出 CSV
  document.getElementById('dlCSV').onclick = () => {
    let csv = "\ufeffSubject,Start Date,Start Time,End Date,End Time,Location,Description\n";
    finalRecords.forEach(r => {
      csv += `"${r.Subject}","${r.DateRaw}","${r.StartT}","${r.DateRaw}","${r.EndT}","${r.Location}","${r.Description}"\n`;
    });
    downloadFile(csv, "university_schedule.csv", "text/csv");
  };

  // 导出 ICS
  document.getElementById('dlICS').onclick = () => {
    let ics = "BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//Gemini Schedule Exporter//CN\nCALSCALE:GREGORIAN\nMETHOD:PUBLISH\n";
    finalRecords.forEach(r => {
      ics += "BEGIN:VEVENT\n";
      ics += `SUMMARY:${r.Subject}\n`;
      ics += `DTSTART:${formatDateICS(r.StartDate, r.StartT)}\n`;
      ics += `DTEND:${formatDateICS(r.StartDate, r.EndT)}\n`;
      ics += `LOCATION:${r.Location}\n`;
      ics += `DESCRIPTION:${r.Description}\n`;
      ics += "STATUS:CONFIRMED\nSEQUENCE:0\nBEGIN:VALARM\nTRIGGER:-PT15M\nACTION:DISPLAY\nDESCRIPTION:Reminder\nEND:VALARM\n";
      ics += "END:VEVENT\n";
    });
    ics += "END:VCALENDAR";
    downloadFile(ics, "university_schedule.ics", "text/calendar");
  };

  function downloadFile(content, fileName, mimeType) {
    const blob = new Blob([content], { type: mimeType + ';charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
  }

  document.getElementById('closeM').onclick = () => container.remove();

})();
