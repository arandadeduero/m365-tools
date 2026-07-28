import { randomInt } from 'node:crypto';
import inquirer from 'inquirer';
import chalk from 'chalk';
import ora from 'ora';
import { getUser, updateUser } from '../graph.js';

import { DOMAIN } from '../constants.js';

// ---------------------------------------------------------------------------
// Password generation — readable Word-NNNN-Word format
// Meets Azure AD complexity: uppercase + lowercase + digit, 12+ chars
// ---------------------------------------------------------------------------

const WORDS = [
  'Acebo', 'Aguila', 'Alamo', 'Almez', 'Aluda', 'Arbol', 'Arco', 'Arena',
  'Arnal', 'Arroyo', 'Avena', 'Azul', 'Barda', 'Bogar', 'Brama', 'Brasa',
  'Bravo', 'Brisa', 'Bronce', 'Buho', 'Cabal', 'Calma', 'Campo', 'Cardo',
  'Ciervo', 'Cima', 'Cisne', 'Claro', 'Cobro', 'Copa', 'Coral', 'Cueva',
  'Dardo', 'Delta', 'Denso', 'Dicho', 'Dique', 'Durce', 'Ebano', 'Encina',
  'Farro', 'Finca', 'Firma', 'Fisco', 'Flama', 'Flora', 'Flujo', 'Fondo',
  'Forja', 'Freno', 'Grano', 'Grifo', 'Gruta', 'Guila', 'Hierro', 'Hondo',
  'Humo', 'Ibero', 'Ingle', 'Islote', 'Jabali', 'Jarco', 'Junco', 'Kilo',
  'Lanza', 'Largo', 'Lauro', 'Lecho', 'Ledge', 'Limon', 'Lince', 'Linde',
  'Lobos', 'Loma', 'Lunar', 'Macho', 'Manto', 'Marco', 'Marga', 'Marte',
  'Metro', 'Mirlo', 'Monte', 'Moral', 'Mosca', 'Nardo', 'Navío', 'Niebla',
  'Norte', 'Nublo', 'Olivo', 'Ombra', 'Onzas', 'Oveja', 'Pardo', 'Pato',
  'Pecho', 'Pedro', 'Peral', 'Perla', 'Piedra', 'Pinar', 'Plata', 'Plaza',
  'Porto', 'Prisa', 'Prosa', 'Purga', 'Radar', 'Ramal', 'Rango', 'Rasgo',
  'Recto', 'Reina', 'Reloj', 'Remo', 'Resto', 'Rioja', 'Risco', 'Roble',
  'Robre', 'Rocio', 'Rodeo', 'Ronda', 'Rubio', 'Rueda', 'Ruido', 'Salmo',
  'Salto', 'Sauce', 'Selva', 'Serra', 'Siena', 'Sierra', 'Solar', 'Sordo',
  'Sucre', 'Surco', 'Tabla', 'Talon', 'Tarde', 'Techo', 'Tejón', 'Termo',
  'Tierra', 'Tigre', 'Timón', 'Tinto', 'Tiro', 'Torca', 'Torno', 'Torre',
  'Tronco', 'Trueno', 'Tupido', 'Umbral', 'Uncio', 'Vados', 'Valle', 'Vasco',
  'Vega', 'Venas', 'Viento', 'Villa', 'Viña', 'Virgo', 'Vista', 'Volco',
  'Yacer', 'Yegua', 'Zarco', 'Zarza', 'Zorro',
];

export function generatePassword() {
  const w1 = WORDS[randomInt(WORDS.length)];
  const w2 = WORDS[randomInt(WORDS.length)];
  const num = String(randomInt(1000, 9999));
  return `${w1}-${num}-${w2}`;
}

// ---------------------------------------------------------------------------
// reset-password command
// ---------------------------------------------------------------------------

/**
 * Reset a user's password and output a mailto: link with credentials.
 * @param {string} upn - Fully resolved UPN (normalizeUpn already applied by index.js)
 */
export async function resetPasswordCommand(upn) {
  // 1. Resolve user
  const spinner = ora(chalk.cyan('Looking up user...')).start();
  const user = await getUser(upn);
  spinner.stop();

  if (!user) {
    console.error(chalk.red(`\nUser not found: ${upn}`));
    process.exit(1);
  }

  const displayName = user.displayName || user.userPrincipalName;
  const userUpn = user.userPrincipalName;

  // 2. Show who we're about to reset
  console.log(chalk.cyan('\n' + '─'.repeat(52)));
  console.log(chalk.bold(`  ${displayName}`));
  console.log(`  ${chalk.gray('UPN:')}        ${userUpn}`);
  if (user.jobTitle) console.log(`  ${chalk.gray('Title:')}      ${user.jobTitle}`);
  if (user.department) console.log(`  ${chalk.gray('Department:')} ${user.department}`);
  console.log(chalk.cyan('─'.repeat(52)));

  // 3. Generate password candidate (shown before confirmation)
  const newPassword = generatePassword();

  console.log(`\n  ${chalk.gray('New password will be:')} ${chalk.bold.green(newPassword)}`);
  console.log(`  ${chalk.gray('(user will be required to change it on first sign-in)')}\n`);

  // 4. Confirmation prompt
  const { confirmed } = await inquirer.prompt([
    {
      type: 'confirm',
      name: 'confirmed',
      message: chalk.yellow(`Reset password for ${chalk.bold(displayName)} (${userUpn})?`),
      default: false,
    },
  ]);

  if (!confirmed) {
    console.log(chalk.gray('\nCancelled. No changes made.\n'));
    return;
  }

  // 5. Apply the password reset via Graph API
  const resetSpinner = ora(chalk.cyan('Applying password reset...')).start();
  try {
    await updateUser(userUpn, {
      passwordProfile: {
        password: newPassword,
        forceChangePasswordNextSignIn: true,
      },
    });
    resetSpinner.succeed(chalk.cyan('Password reset applied.'));
  } catch (err) {
    resetSpinner.fail(chalk.red(`\nFailed to reset password: ${err.message || err}`));
    process.exit(1);
  }

  // 6. Build mailto: link
  // Body text (plain, will be URL-encoded):
  //   username: upn@arandadeduero.es
  //   password: <generated>
  //   url: https://aytoarandaduero.sharepoint.com/
  const body = [
    `username: ${userUpn}`,
    `password: ${newPassword}`,
    `Recuerda que la contraseña de Microsoft365 es diferente a la de tu actual correo corporativo y a la de tu ordenador. Esta contraseña es temporal y deberás cambiarla en tu primer inicio de sesión.`,
    `\n`,
    `url: https://aytoarandaduero.sharepoint.com/`,
    `\n`,
    `Sigue la guía de bienvenida en este enlace https://aytoarandaduero-my.sharepoint.com/:f:/g/personal/glopez_${DOMAIN.replace('.', '_')}/IgDxAF3YuiHuRILrP8vCkGRnAYrtXe0_jzxtD_l9ZON1CGg?e=86PJVY`,
  ].join('\n');

  const mailtoLink =
    `mailto:${encodeURIComponent(userUpn)}` +
    `?subject=${encodeURIComponent('Microsoft365')}` +
    `&body=${encodeURIComponent(body)}`;

  // 7. Print result
  console.log('\n' + chalk.green('✓') + chalk.bold(` Password reset for ${displayName}`));
  console.log(chalk.cyan('─'.repeat(52)));
  console.log(`  ${chalk.gray('UPN:')}      ${userUpn}`);
  console.log(`  ${chalk.gray('Password:')} ${chalk.bold(newPassword)}`);
  console.log(chalk.cyan('─'.repeat(52)));
  console.log(`\n  ${chalk.gray('mailto link (open in your email client):')}`);
  console.log(`\n  ${chalk.blue(mailtoLink)}\n`);
}
