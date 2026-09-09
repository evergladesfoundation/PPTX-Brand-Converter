/// <reference types="astro/client" />

declare namespace App {
  interface Locals {
    authRequired: boolean;
    authed: boolean;
  }
}
