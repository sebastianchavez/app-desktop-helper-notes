import { bootstrapApplication } from '@angular/platform-browser';
import { appConfig } from './app/app.config';
import { App } from './app/app';

// Punto de entrada del renderer: crea la aplicación standalone con sus providers globales.
bootstrapApplication(App, appConfig)
  .catch((err) => console.error(err));
