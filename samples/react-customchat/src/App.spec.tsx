import { test, expect } from '@playwright/experimental-ct-react17';
import App from './App';

test.use({ viewport: { width: 500, height: 500 } });

test('should work', async ({ mount }) => {
  const component = await mount(<App />);
  await expect(component).toContainText('Learn React');
});

test('has title', async ({ page }) => {
    await page.goto('https://www.bing.com/');
    // Expect a title "to contain" a substring => OpenReplay
    await expect(page).toHaveTitle("Bing");
});

   
   