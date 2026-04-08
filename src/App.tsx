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
  MessageBarType,
  Spinner,
  SpinnerSize,
  Link
} from '@fluentui/react';
import { convertSelection, convertMainBody, flattenAdvancedElements } from './converter';
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
  const [isProcessing, setIsProcessing] = useState(false);
  const [isOffice, setIsOffice] = useState(true);

  useState(() => {
    // Check if we are running inside Office
    if (typeof Office !== 'undefined' && !Office.context?.host) {
      setIsOffice(false);
    }
  });

  const handleAction = async (action: () => Promise<void>) => {
    try {
      setStatus(null);
      setIsProcessing(true);
      await action();
      setStatus({ message: "ดำเนินการเรียบร้อย!", type: MessageBarType.success });
    } catch (error: any) {
      console.error(error);
      setStatus({ message: `Error: ${error.message || "ล้มเหลว"}`, type: MessageBarType.error });
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <ThemeProvider theme={myTheme}>
      <Stack styles={stackItemPadding} tokens={stackSpacing}>
        <Stack horizontal verticalAlign="center">
          <Text styles={logoStyle}>IT๙</Text>
          <Stack>
            <Text variant="xxLarge" styles={boldStyle}>Converter</Text>
            <Text variant="small" styles={{ root: { color: '#605e5c', marginTop: '-5px' } }}>v1.22.0</Text>
          </Stack>
        </Stack>
        
        <Text variant="medium">
          เครื่องมือแปลงเลขไทย พร้อมระบบข้ามคำภาษาอังกฤษและ URL อัตโนมัติ
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
          disabled={isProcessing}
        />
        <Text variant="small" styles={{ root: { color: '#666', marginTop: '-15px', marginLeft: '25px' } }}>
          ข้ามเลขในคำอังกฤษ เช่น spin9, 9arm, www.site123.com
        </Text>

        <Separator />

        <Stack tokens={{ childrenGap: 10 }}>
          <PrimaryButton 
            text="แปลงเนื้อหาหลัก" 
            onClick={() => handleAction(() => convertMainBody(useSmartIgnore))} 
            iconProps={{ iconName: 'Document' }}
            disabled={isProcessing}
          />
          <PrimaryButton 
            text="แปลงรวมเลขหน้าและเลขหัวข้อแบบ Flatten" 
            onClick={() => handleAction(() => flattenAdvancedElements(useSmartIgnore))} 
            iconProps={{ iconName: 'FullWidth' }}
            disabled={isProcessing}
          />
          <DefaultButton 
            text="แปลงเฉพาะที่เลือก" 
            onClick={() => handleAction(() => convertSelection(useSmartIgnore))} 
            iconProps={{ iconName: 'SingleColumn' }}
            disabled={isProcessing}
          />
        </Stack>

        {isProcessing && (
          <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 10 }}>
            <Spinner size={SpinnerSize.medium} />
            <Text>กำลังดำเนินการ...</Text>
          </Stack>
        )}

        {status && !isProcessing && (
          <MessageBar
            messageBarType={status.type}
            isMultiline={true}
            onDismiss={() => setStatus(null)}
          >
            {status.message}
          </MessageBar>
        )}

        {!isOffice && (
          <Stack tokens={{ childrenGap: 10 }} styles={{ root: { backgroundColor: '#f3f2f1', padding: '15px', borderRadius: '4px', border: '1px dashed #d2d0ce', marginTop: '10px' } }}>
            <Text variant="small" styles={{ root: { fontWeight: FontWeights.semibold, textAlign: 'center' } }}>
              IT๙ Converter เป็นส่วนเสริมสำหรับ Microsoft Word
            </Text>
            <DefaultButton 
              text="ดูวิธีการติดตั้งบน GitHub" 
              iconProps={{ iconName: 'ReadingMode' }}
              href="https://github.com/vk-poonyakanok/thai-numeral-converter#วิธีการใช้งานสำหรับ-macos-how-to-use"
              target="_blank"
              styles={{ 
                root: { 
                  backgroundColor: '#fff4ce', 
                  borderColor: '#fde79e',
                  color: '#323130'
                },
                rootHovered: {
                  backgroundColor: '#fde79e',
                  borderColor: '#fcdb6d'
                }
              }}
            />
            <Text variant="xSmall" styles={{ root: { textAlign: 'center', color: '#605e5c' } }}>
              เครื่องมือนี้ออกแบบมาเพื่อใช้งานใน Word กรุณาดูวิธีติดตั้งเพื่อเริ่มต้นใช้งาน
            </Text>
          </Stack>
        )}

        <Separator />
        
        <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 15 }}>
          <Link href="https://vk-poonyakanok.github.io/thai-numeral-converter/privacy.html" target="_blank" styles={{ root: { fontSize: '10px' } }}>
            Privacy Policy
          </Link>
          <Link href="https://vk-poonyakanok.github.io/thai-numeral-converter/eula.html" target="_blank" styles={{ root: { fontSize: '10px' } }}>
            Terms of Use
          </Link>
        </Stack>

        <Text variant="small" styles={{ root: { textAlign: 'center', color: '#a19f9d', display: 'block', marginTop: '5px' } }}>
          IT๙ Converter v1.22.0
        </Text>
      </Stack>
    </ThemeProvider>
  );
}

export default App;
