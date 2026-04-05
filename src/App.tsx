import { useState } from 'react';
import { 
  PrimaryButton, 
  DefaultButton, 
  Toggle, 
  Stack, 
  Text, 
  ThemeProvider, 
  createTheme, 
  FontWeights, 
  Separator,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import { convertSelection, convertDocument } from './converter';
import './App.css';

const myTheme = createTheme({
  palette: {
    themePrimary: '#d83b01',
    themeLighterAlt: '#fdf6f2',
    themeLighter: '#f7dbcf',
    themeLight: '#f0beaa',
    themeTertiary: '#e18360',
    themeSecondary: '#dc5120',
    themeDarkAlt: '#c23501',
    themeDark: '#a42d01',
    themeDarker: '#792101',
    neutralLighterAlt: '#f8f8f8',
    neutralLighter: '#f4f4f4',
    neutralLight: '#eaeaea',
    neutralQuaternaryAlt: '#dadada',
    neutralQuaternary: '#d0d0d0',
    neutralTertiaryAlt: '#c8c8c8',
    neutralTertiary: '#a19f9d',
    neutralSecondary: '#605e5c',
    neutralPrimaryAlt: '#3b3a39',
    neutralPrimary: '#323130',
    neutralDark: '#201f1e',
    black: '#000000',
    white: '#ffffff',
  }
});

const boldStyle = { root: { fontWeight: FontWeights.semibold } };
const logoStyle = { 
  root: { 
    background: '#d83b01', 
    color: 'white', 
    fontWeight: 'bold', 
    padding: '4px 8px', 
    borderRadius: '4px',
    fontSize: '24px',
    marginRight: '12px'
  } 
};
const stackItemPadding = { root: { padding: 15 } };
const stackSpacing = { childrenGap: 15 };

function App() {
  const [useSmartIgnore, setUseSmartIgnore] = useState(true);
  const [status, setStatus] = useState<{message: string, type: MessageBarType} | null>(null);

  const handleConvertSelection = async () => {
    try {
      setStatus(null);
      await convertSelection(useSmartIgnore);
      setStatus({ message: "แปลงเฉพาะที่เลือกเรียบร้อย!", type: MessageBarType.success });
    } catch (error: any) {
      console.error(error);
      setStatus({ message: `Error: ${error.message || "ล้มเหลว"}`, type: MessageBarType.error });
    }
  };

  const handleConvertDocument = async () => {
    try {
      setStatus(null);
      await convertDocument(useSmartIgnore);
      setStatus({ message: "แปลงทั้งเอกสารเรียบร้อย!", type: MessageBarType.success });
    } catch (error: any) {
      console.error(error);
      setStatus({ message: `Error: ${error.message || "ล้มเหลว"}`, type: MessageBarType.error });
    }
  };

  return (
    <ThemeProvider theme={myTheme}>
      <Stack styles={stackItemPadding} tokens={stackSpacing}>
        <Stack horizontal verticalAlign="center">
          <Text styles={logoStyle}>IT๙</Text>
          <Stack>
            <Text variant="xxLarge" styles={boldStyle}>IT๙ Converter</Text>
            <Text variant="small" styles={{ root: { color: '#605e5c', marginTop: '-5px' } }}>v1.15.0</Text>
          </Stack>
        </Stack>
        
        <Text variant="medium">
          เครื่องมือแปลงเลขไทยระดับมืออาชีพ พร้อมระบบข้ามคำภาษาอังกฤษและ URL อัตโนมัติ (แปลงครอบคลุมทั้งหัวกระดาษ ท้ายกระดาษ กล่องข้อความ เลขหน้า และลำดับอัตโนมัติ)
        </Text>

        <Separator />

        <Text variant="medium" styles={boldStyle}>การตั้งค่า (Options)</Text>
        
        <Toggle 
          label="Smart Ignore" 
          inlineLabel 
          checked={useSmartIgnore} 
          onChange={(_, checked) => setUseSmartIgnore(!!checked)}
          onText="On"
          offText="Off"
        />
        <Text variant="small" styles={{ root: { color: '#666', marginTop: '-15px', marginLeft: '25px' } }}>
          ข้ามเลขในคำอังกฤษ เช่น spin9, 9arm, www.site123.com
        </Text>

        <Separator />

        <Stack tokens={{ childrenGap: 10 }}>
          <PrimaryButton 
            text="แปลงทั้งเอกสาร" 
            onClick={handleConvertDocument} 
            iconProps={{ iconName: 'Document' }}
          />
          <DefaultButton 
            text="แปลงเฉพาะที่เลือก" 
            onClick={handleConvertSelection} 
            iconProps={{ iconName: 'SingleColumn' }}
          />
        </Stack>

        {status && (
          <MessageBar
            messageBarType={status.type}
            isMultiline={true}
            onDismiss={() => setStatus(null)}
          >
            {status.message}
          </MessageBar>
        )}

        <Separator />
        
        <Text variant="small" styles={{ root: { textAlign: 'center', color: '#a19f9d' } }}>
          Professional Arabic to Thai Numeral Tool
        </Text>
      </Stack>
    </ThemeProvider>
  );
}

export default App;
