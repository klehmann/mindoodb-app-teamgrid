import { createApp } from "vue";
import PrimeVue from "primevue/config";
import Tooltip from "primevue/tooltip";
import "@fontsource/inter/latin-400.css";
import "@fontsource/inter/latin-600.css";
import "@fontsource/inter/latin-700.css";
import "@fontsource/rubik/latin-400.css";
import "@fontsource/rubik/latin-700.css";
import "@fontsource/jetbrains-mono/latin-400.css";
import "@fontsource/carlito/latin-400.css";
import "@fontsource/carlito/latin-700.css";
import "@fontsource/caladea/latin-400.css";
import "@fontsource/caladea/latin-700.css";
import "@fontsource/cousine/latin-400.css";
import "@fontsource/cousine/latin-700.css";
import "@fontsource/arimo/latin-400.css";
import "@fontsource/arimo/latin-700.css";
import "@fontsource/tinos/latin-400.css";
import "@fontsource/tinos/latin-700.css";
import "primeicons/primeicons.css";

import App from "./App.vue";
import "@/assets/styles/main.css";
import { applyAppTheme, buildPrimeVueTheme } from "@/shared/lib/theme";

const app = createApp(App);

app.use(PrimeVue, {
  ripple: true,
  theme: buildPrimeVueTheme(),
});
app.directive("tooltip", Tooltip);

applyAppTheme();
app.mount("#app");
