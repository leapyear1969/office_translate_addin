import { UserProfile } from './api';

export function setupAccount(onSignOut: () => void) {
  const avatar = document.getElementById('account-avatar') as HTMLButtonElement;
  const menu = document.getElementById('account-menu')!;
  const photo = document.getElementById('account-photo') as HTMLImageElement;
  const initial = document.getElementById('account-initial')!;
  const signout = document.getElementById('signout') as HTMLButtonElement;
  function close(restoreFocus = false) {
    menu.hidden = true;
    avatar.setAttribute('aria-expanded', 'false');
    if (restoreFocus) avatar.focus();
  }
  avatar.addEventListener('click', () => {
    if (!menu.hidden) { close(); return; }
    menu.hidden = false;
    avatar.setAttribute('aria-expanded', 'true');
    menu.querySelector<HTMLButtonElement>('button:not([hidden]):not(:disabled)')?.focus();
  });
  document.addEventListener('click', event => {
    if (!avatar.parentElement!.contains(event.target as Node)) close();
  });
  document.addEventListener('keydown', event => {
    if (event.key === 'Escape' && !menu.hidden) { event.preventDefault(); close(true); }
  });
  avatar.parentElement!.addEventListener('focusout', event => {
    if (event.relatedTarget && !avatar.parentElement!.contains(event.relatedTarget as Node)) close();
  });
  photo.addEventListener('error', () => { photo.hidden = true; initial.hidden = false; });
  signout.addEventListener('click', () => { onSignOut(); close(true); });
  return {
    close,
    render(user?: UserProfile) {
      const name = user?.displayName || user?.mail || '';
      initial.textContent = Array.from(name)[0]?.toUpperCase() || '人';
      initial.hidden = !!user?.photo;
      photo.hidden = !user?.photo;
      if (user?.photo) photo.src = user.photo;
      else photo.removeAttribute('src');
      avatar.setAttribute('aria-label', name ? `${name}，用户账户` : '用户账户，未登录');
      avatar.title = name || '用户账户';
      signout.disabled = !user;
    },
  };
}
