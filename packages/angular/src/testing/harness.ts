// Helpers for the Angular specs. Not part of the package build.
import { provideZonelessChangeDetection, type Provider, type Type } from '@angular/core';
import { TestBed, type ComponentFixture } from '@angular/core/testing';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';

/** Configure the test module for a zoneless app, with extra providers. */
export function configureZoneless(providers: Provider[] = []): void {
  TestBed.configureTestingModule({
    providers: [provideZonelessChangeDetection(), ...providers],
  });
}

/** Wait for Angular, the engine's deferred notifications, and its layout. */
export async function settle(fixture?: ComponentFixture<unknown>, rounds = 4): Promise<void> {
  for (let round = 0; round < rounds; round += 1) {
    await fixture?.whenStable();
    await new Promise((resolve) => setTimeout(resolve, 20));
  }
  await fixture?.whenStable();
}

/** Create a host component, render it, and wait until its first editor has a painted page. */
export async function render<T>(component: Type<T>): Promise<ComponentFixture<T>> {
  const fixture = TestBed.createComponent(component);
  await settle(fixture);
  return fixture;
}

/** The engine of a root, or a failure when it has none. */
export function engineOf(editor: DocxEditorInstance | null | undefined): DocxEditorInstance {
  if (!editor) throw new Error('the root has no engine');
  return editor;
}

/** Put the caret at the end of the first paragraph and type `text`. */
export function typeText(editor: DocxEditorInstance, text: string): void {
  const surface = editor.surface;
  if (!surface) throw new Error('the engine is not attached');
  const paragraphId = surface.session.paragraphIds()[0]!;
  surface.setSelection({ anchor: { paragraphId, offset: 0 }, head: { paragraphId, offset: 0 } });
  const result = editor.exec({ type: 'insertText', text });
  if (!result.ok) throw new Error(`the engine refused the text: ${result.reason}`);
}
