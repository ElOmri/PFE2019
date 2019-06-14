import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Lists } from 'sp-pnp-js';

export interface IPnPGraphProps {
  description: string,
  context: WebPartContext,
  Lists: string,
  TemplateFile: string;
}
