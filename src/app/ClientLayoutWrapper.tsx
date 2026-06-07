'use client'

import StudioLayout from '@/components/StudioLayout';

export default function ClientLayoutWrapper({
  children,
}: {
  children: React.ReactNode;
}) {
  return <StudioLayout>{children}</StudioLayout>;
}
