import { createApp } from 'vue'
import { createPinia } from 'pinia'
import PrimeVue from 'primevue/config'
import DataTable from 'primevue/datatable'
import Column from 'primevue/column'
import InputText from 'primevue/inputtext'

// Import PrimeVue styles
import 'primevue/resources/primevue.min.css'
import 'primeicons/primeicons.css'

// Import a PrimeVue theme (you can choose a different theme if you prefer)
import 'primevue/resources/themes/lara-light-blue/theme.css'

import App from './App.vue'
import router from './router'

const app = createApp(App)

app.use(createPinia())
app.use(router)
app.use(PrimeVue)

app.component('DataTable', DataTable)
app.component('Column', Column)
app.component('InputText', InputText)

app.mount('#app')