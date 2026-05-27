import { createRoot } from 'react-dom/client';
import '@mog-sdk/spreadsheet-app/styles.css';
import './styles.css';
import { App } from './App';

createRoot(document.getElementById('root')!).render(<App />);
