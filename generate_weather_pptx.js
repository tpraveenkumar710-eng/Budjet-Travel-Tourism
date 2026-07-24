import pptxgen from 'pptxgenjs';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const pptx = new pptxgen();

// Presentation Layout & Properties
pptx.layout = 'LAYOUT_16x9';
pptx.title = 'Weather Forecasting UI & Smart Travel Intelligence';
pptx.author = 'ExploreBharat UI/UX Engineering Team';
pptx.company = 'ExploreBharat Capstone Project';

// Premium Color Palette
const BG_DARK = '090D16';
const CARD_BG = '1E293B';
const PRIMARY = '6366F1';
const SECONDARY = '38BDF8';
const ACCENT = 'C084FC';
const TEXT_WHITE = 'F8FAFC';
const TEXT_MUTED = '94A3B8';
const SUCCESS = '34D399';
const WARNING = 'F59E0B';

// Define Master Slide
pptx.defineSlideMaster({
  title: 'WEATHER_MASTER',
  background: { color: BG_DARK },
  objects: [
    { rect: { x: 0, y: 0, w: '100%', h: 0.15, fill: { color: SECONDARY } } },
    { text: { text: 'ExploreBharat — Weather Forecasting UI & Climate Intelligence PPT', options: { x: 0.5, y: 7.1, w: 7, h: 0.3, fontSize: 10, color: TEXT_MUTED } } }
  ]
});

// SLIDE 1: Title Slide
const slide1 = pptx.addSlide();
slide1.background = { color: BG_DARK };
slide1.addText('WEATHER FORECASTING UI', {
  x: 0.8, y: 2.0, w: 11.5, h: 1.0,
  fontSize: 44, bold: true, color: SECONDARY, fontFace: 'Trebuchet MS', align: 'left'
});
slide1.addText('Real-Time Climate Intelligence & Automated Travel Safety Advisory System', {
  x: 0.8, y: 3.1, w: 11.5, h: 0.8,
  fontSize: 22, color: TEXT_WHITE, fontFace: 'Calibri'
});
slide1.addText('Integrated into ExploreBharat Budget Travel Engine • React 19 • Glassmorphism UI • REST APIs', {
  x: 0.8, y: 4.5, w: 11.5, h: 0.5,
  fontSize: 14, color: ACCENT, fontFace: 'Calibri'
});

// SLIDE 2: Problem Statement & Necessity
const slide2 = pptx.addSlide({ masterName: 'WEATHER_MASTER' });
slide2.addText('Problem & Market Need', { x: 0.6, y: 0.4, w: 10, h: 0.6, fontSize: 26, bold: true, color: SECONDARY });
slide2.addText('Why Weather Forecasting is Crucial for Budget Travelers', { x: 0.6, y: 1.0, w: 10, h: 0.4, fontSize: 14, color: TEXT_MUTED });

slide2.addShape(pptx.shapes.RECTANGLE, { x: 0.6, y: 1.6, w: 3.8, h: 5.0, fill: { color: CARD_BG }, line: { color: PRIMARY, width: 1 } });
slide2.addText('Unexpected Weather Risks', { x: 0.8, y: 1.9, w: 3.4, h: 0.4, fontSize: 18, bold: true, color: WARNING });
slide2.addText('• Sudden mountain rain in Munnar/Ooty ruins trekking plans.\n• Extreme heatwaves in Jaipur affect outdoor sightseeing.\n• High ocean tides in Goa affect beach safety.', { x: 0.8, y: 2.5, w: 3.4, h: 3.8, fontSize: 13, color: TEXT_WHITE });

slide2.addShape(pptx.shapes.RECTANGLE, { x: 4.7, y: 1.6, w: 3.8, h: 5.0, fill: { color: CARD_BG }, line: { color: SECONDARY, width: 1 } });
slide2.addText('Budget Impact', { x: 4.9, y: 1.9, w: 3.4, h: 0.4, fontSize: 18, bold: true, color: SECONDARY });
slide2.addText('• Rain delays cause wasted hotel bookings.\n• Emergency transport re-routing increases trip costs.\n• Lack of advisory leads to missed activity passes.', { x: 4.9, y: 2.5, w: 3.4, h: 3.8, fontSize: 13, color: TEXT_WHITE });

slide2.addShape(pptx.shapes.RECTANGLE, { x: 8.8, y: 1.6, w: 3.8, h: 5.0, fill: { color: CARD_BG }, line: { color: SUCCESS, width: 1 } });
slide2.addText('Our Smart Solution', { x: 9.0, y: 1.9, w: 3.4, h: 0.4, fontSize: 18, bold: true, color: SUCCESS });
slide2.addText('• Live 5-Day forecast per destination.\n• Automated safety advisory engine.\n• Dynamic °C/°F unit conversion.\n• Direct integration with budget calculator.', { x: 9.0, y: 2.5, w: 3.4, h: 3.8, fontSize: 13, color: TEXT_WHITE });

// SLIDE 3: Key UI/UX Features
const slide3 = pptx.addSlide({ masterName: 'WEATHER_MASTER' });
slide3.addText('Key UI/UX Features', { x: 0.6, y: 0.4, w: 10, h: 0.6, fontSize: 26, bold: true, color: SECONDARY });
slide3.addText('Interactive Design Elements Built for Seamless User Experience', { x: 0.6, y: 1.0, w: 10, h: 0.4, fontSize: 14, color: TEXT_MUTED });

const features = [
  { title: 'Glassmorphism Cards', desc: 'Sleek dark backdrop with translucent frosted glass card containers and subtle glowing borders.' },
  { title: '5-Day Forecast Grid', desc: 'Daily weather cards showcasing day, weather icons, rain probability %, and temperature trends.' },
  { title: 'Weather Metrics', desc: 'Real-time telemetry showing Humidity %, Wind Speed (km/h), UV Index, and Feel Temperature.' },
  { title: 'Smart Advisory Banner', desc: 'Contextual travel advisory warning users about weather conditions before booking.' }
];

features.forEach((f, idx) => {
  const xPos = idx % 2 === 0 ? 0.6 : 6.8;
  const yPos = idx < 2 ? 1.7 : 4.4;
  slide3.addShape(pptx.shapes.RECTANGLE, { x: xPos, y: yPos, w: 5.8, h: 2.4, fill: { color: CARD_BG }, line: { color: ACCENT, width: 1 } });
  slide3.addText(f.title, { x: xPos + 0.3, y: yPos + 0.3, w: 5.2, h: 0.4, fontSize: 18, bold: true, color: SECONDARY });
  slide3.addText(f.desc, { x: xPos + 0.3, y: yPos + 0.8, w: 5.2, h: 1.3, fontSize: 13, color: TEXT_WHITE });
});

// SLIDE 4: Technical Architecture & Data Flow
const slide4 = pptx.addSlide({ masterName: 'WEATHER_MASTER' });
slide4.addText('Technical Architecture & Data Pipeline', { x: 0.6, y: 0.4, w: 10, h: 0.6, fontSize: 26, bold: true, color: SECONDARY });
slide4.addText('Seamless Integration between Frontend State & Weather Telemetry Data Engine', { x: 0.6, y: 1.0, w: 10, h: 0.4, fontSize: 14, color: TEXT_MUTED });

slide4.addShape(pptx.shapes.RECTANGLE, { x: 0.6, y: 1.8, w: 3.5, h: 4.8, fill: { color: CARD_BG }, line: { color: PRIMARY, width: 1 } });
slide4.addText('1. User Interaction', { x: 0.8, y: 2.1, w: 3.1, h: 0.4, fontSize: 16, bold: true, color: PRIMARY });
slide4.addText('• User selects destination (e.g. Goa, Ooty, Manali).\n• Toggle °C / °F temperature unit.\n• Real-time reactive state updates UI instantly.', { x: 0.8, y: 2.7, w: 3.1, h: 3.5, fontSize: 13, color: TEXT_WHITE });

slide4.addShape(pptx.shapes.RECTANGLE, { x: 4.8, y: 1.8, w: 3.5, h: 4.8, fill: { color: CARD_BG }, line: { color: SECONDARY, width: 1 } });
slide4.addText('2. Weather Engine', { x: 5.0, y: 2.1, w: 3.1, h: 0.4, fontSize: 16, bold: true, color: SECONDARY });
slide4.addText('• Fetches telemetry dataset (Temp, Wind, Humidity, UV).\n• Computes 5-day rain probability & icon mapping.\n• Generates contextual travel safety score.', { x: 5.0, y: 2.7, w: 3.1, h: 3.5, fontSize: 13, color: TEXT_WHITE });

slide4.addShape(pptx.shapes.RECTANGLE, { x: 9.0, y: 1.8, w: 3.5, h: 4.8, fill: { color: CARD_BG }, line: { color: SUCCESS, width: 1 } });
slide4.addText('3. Budget Integration', { x: 9.2, y: 2.1, w: 3.1, h: 0.4, fontSize: 16, bold: true, color: SUCCESS });
slide4.addText('• Passes target destination to Budget Calculator.\n• Recommends indoor vs outdoor activity packages.\n• Ensures safe & cost-effective trip planning.', { x: 9.2, y: 2.7, w: 3.1, h: 3.5, fontSize: 13, color: TEXT_WHITE });

// SLIDE 5: Technology Stack
const slide5 = pptx.addSlide({ masterName: 'WEATHER_MASTER' });
slide5.addText('Technology Stack & Libraries Used', { x: 0.6, y: 0.4, w: 10, h: 0.6, fontSize: 26, bold: true, color: SECONDARY });

const stack = [
  { name: 'React 19 & Hooks', detail: 'State management using useState, useEffect, and component modularity.' },
  { name: 'Lucide React Icons', detail: 'Dynamic weather icons (Sun, CloudSun, CloudRain, Snowflake, Wind, Droplets).' },
  { name: 'Vanilla CSS Design', detail: 'Custom glassmorphic design tokens, glowing flex containers, and responsive grids.' },
  { name: 'Node & Express API', detail: 'Backend REST API endpoint server supporting single unified deployment.' }
];

stack.forEach((s, idx) => {
  const y = 1.6 + idx * 1.3;
  slide5.addShape(pptx.shapes.RECTANGLE, { x: 0.6, y, w: 12.0, h: 1.1, fill: { color: CARD_BG }, line: { color: SECONDARY, width: 1 } });
  slide5.addText(s.name, { x: 0.9, y: y + 0.2, w: 3.5, h: 0.4, fontSize: 16, bold: true, color: ACCENT });
  slide5.addText(s.detail, { x: 4.5, y: y + 0.2, w: 7.8, h: 0.6, fontSize: 13, color: TEXT_WHITE });
});

// SLIDE 6: Summary & Impact
const slide6 = pptx.addSlide({ masterName: 'WEATHER_MASTER' });
slide6.addText('Summary & Value Addition', { x: 0.6, y: 0.4, w: 10, h: 0.6, fontSize: 26, bold: true, color: SECONDARY });

slide6.addShape(pptx.shapes.RECTANGLE, { x: 0.6, y: 1.6, w: 12.0, h: 5.0, fill: { color: CARD_BG }, line: { color: SUCCESS, width: 1 } });
slide6.addText('Key Takeaways & Benefits:', { x: 1.0, y: 2.0, w: 11.2, h: 0.5, fontSize: 20, bold: true, color: SUCCESS });
slide6.addText('1. Empowered Travelers: Provides instant visibility into climate conditions across India.\n2. Risk Mitigation: Prevents wasted bookings by highlighting rain & extreme temperature advisories.\n3. Seamless UX: Fast, reactive, responsive interface built into ExploreBharat.\n4. Scalable Infrastructure: Ready for live OpenWeatherMap API integration.', { x: 1.0, y: 2.7, w: 11.2, h: 3.5, fontSize: 15, color: TEXT_WHITE, lineSpacing: 24 });

// Save PPTX presentation
const outputPath = path.join(__dirname, 'Weather_Forecasting_UI_Presentation.pptx');
const publicPath = path.join(__dirname, 'public', 'Weather_Forecasting_UI_Presentation.pptx');

pptx.writeFile({ fileName: outputPath }).then(() => {
  fs.copyFileSync(outputPath, publicPath);
  console.log(`✅ Weather Forecasting UI Presentation PPTX generated successfully at:\n  - ${outputPath}\n  - ${publicPath}`);
}).catch(err => {
  console.error('Error generating Weather PPTX:', err);
});
