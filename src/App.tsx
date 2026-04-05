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
const stackItemPadding = { root: { padding: 10 } };
const stackSpacing = { childrenGap: 20 };

function App() {
  const [useSmartIgnore, setUseSmartIgnore] = useState(true);
  const [status, setStatus] = useState<{message: string, type: MessageBarType} | null>(null);

  const handleConvertSelection = async () => {
    try {
      setStatus(null);
      await convertSelection(useSmartIgnore);
      setStatus({ message: "Selection converted successfully!", type: MessageBarType.success });
    } catch (error: any) {
      console.error(error);
      setStatus({ message: `Error: ${error.message || "Failed to convert selection"}`, type: MessageBarType.error });
    }
  };

  const handleConvertDocument = async () => {
    try {
      setStatus(null);
      await convertDocument(useSmartIgnore);
      setStatus({ message: "Entire document converted successfully!", type: MessageBarType.success });
    } catch (error: any) {
      console.error(error);
      setStatus({ message: `Error: ${error.message || "Failed to convert document"}`, type: MessageBarType.error });
    }
  };

  return (
    <ThemeProvider theme={myTheme}>
      <Stack styles={stackItemPadding} tokens={stackSpacing}>
        <Stack horizontal verticalAlign="center">
          <Text variant="xxLarge" styles={boldStyle}>Thai Numeral Converter</Text>
        </Stack>
        
        <Text variant="medium">
          Easily convert Arabic numerals (0-9) to Thai numerals (๐-๙).
        </Text>

        <Separator />

        <Toggle 
          label="Smart Ignore" 
          inlineLabel 
          checked={useSmartIgnore} 
          onChange={(_, checked) => setUseSmartIgnore(!!checked)}
          onText="On"
          offText="Off"
        />
        <Text variant="small" styles={{ root: { color: '#666', marginTop: '-15px' } }}>
          Safely skip numbers embedded in English words (e.g. spin9, 9arm), URLs, and emails.
        </Text>

        <Stack tokens={{ childrenGap: 10 }}>
          <PrimaryButton 
            text="Convert Entire Document" 
            onClick={handleConvertDocument} 
            iconProps={{ iconName: 'Document' }}
          />
          <DefaultButton 
            text="Convert Selection" 
            onClick={handleConvertSelection} 
            iconProps={{ iconName: 'SingleColumn' }}
          />
        </Stack>

        {status && (
          <MessageBar
            messageBarType={status.type}
            isMultiline={false}
            onDismiss={() => setStatus(null)}
            dismissButtonAriaLabel="Close"
          >
            {status.message}
          </MessageBar>
        )}

        <Separator />
        
        <Text variant="small">
          Built for Microsoft Word Web Add-in
        </Text>
      </Stack>
    </ThemeProvider>
  );
}

export default App;
