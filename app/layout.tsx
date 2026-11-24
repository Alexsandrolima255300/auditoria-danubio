import './globals.css';

export const metadata = { title: 'ALEX FREIGHT AUDIT', description: 'Auditoria de fretes Danúbio' };

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="pt-BR"><body>{children}</body></html>;
}