"use client";

import { useEffect, useState } from "react";
import { KeyRound, UserRound } from "lucide-react";

import { useMeetFlow } from "@/components/meetflow/app-context";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";

export default function AccountPage() {
  const { userName, userEmail, updateProfile, changePassword, refreshProfile } = useMeetFlow();
  const [name, setName] = useState(userName || "");
  const [currentPassword, setCurrentPassword] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [profileStatus, setProfileStatus] = useState("");
  const [passwordStatus, setPasswordStatus] = useState("");
  const [savingProfile, setSavingProfile] = useState(false);
  const [savingPassword, setSavingPassword] = useState(false);

  useEffect(() => {
    setName(userName || "");
  }, [userName]);

  useEffect(() => {
    refreshProfile().catch(() => {});
  }, []);

  async function onSaveProfile(e) {
    e.preventDefault();
    setProfileStatus("");
    setSavingProfile(true);
    try {
      await updateProfile({ full_name: name });
      setProfileStatus("Nome atualizado com sucesso.");
    } catch (err) {
      setProfileStatus(err?.message || "Não foi possível atualizar seu nome.");
    } finally {
      setSavingProfile(false);
    }
  }

  async function onChangePassword(e) {
    e.preventDefault();
    setPasswordStatus("");
    if (newPassword !== confirmPassword) {
      setPasswordStatus("A nova senha e a confirmação não conferem.");
      return;
    }
    setSavingPassword(true);
    try {
      await changePassword({ current_password: currentPassword, new_password: newPassword });
      setCurrentPassword("");
      setNewPassword("");
      setConfirmPassword("");
      setPasswordStatus("Senha alterada com sucesso.");
    } catch (err) {
      setPasswordStatus(err?.message || "Não foi possível alterar sua senha.");
    } finally {
      setSavingPassword(false);
    }
  }

  return (
    <div className="space-y-6">
      <Card>
        <CardHeader>
          <CardTitle>Conta</CardTitle>
          <CardDescription>Gerencie suas informações e segurança de acesso.</CardDescription>
        </CardHeader>
      </Card>

      <div className="grid gap-6 xl:grid-cols-2">
        <Card>
          <CardHeader>
            <div className="flex items-center gap-2">
              <UserRound className="h-4 w-4 text-blue-300" />
              <CardTitle>Perfil</CardTitle>
            </div>
            <CardDescription>Defina o nome que será exibido no workspace.</CardDescription>
          </CardHeader>
          <CardContent>
            <form className="space-y-3" onSubmit={onSaveProfile}>
              <Input value={name} onChange={(e) => setName(e.target.value)} placeholder="Seu nome" />
              <Input value={userEmail || ""} disabled />
              {profileStatus ? <p className="text-sm text-slate-300">{profileStatus}</p> : null}
              <Button type="submit" disabled={savingProfile}>
                {savingProfile ? "Salvando..." : "Salvar perfil"}
              </Button>
            </form>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <div className="flex items-center gap-2">
              <KeyRound className="h-4 w-4 text-lime-300" />
              <CardTitle>Senha</CardTitle>
            </div>
            <CardDescription>Altere sua senha com segurança.</CardDescription>
          </CardHeader>
          <CardContent>
            <form className="space-y-3" onSubmit={onChangePassword}>
              <Input
                type="password"
                value={currentPassword}
                onChange={(e) => setCurrentPassword(e.target.value)}
                placeholder="Senha atual"
              />
              <Input
                type="password"
                value={newPassword}
                onChange={(e) => setNewPassword(e.target.value)}
                placeholder="Nova senha"
              />
              <Input
                type="password"
                value={confirmPassword}
                onChange={(e) => setConfirmPassword(e.target.value)}
                placeholder="Confirmar nova senha"
              />
              <p className="text-xs text-slate-500">
                Use no mínimo 8 caracteres, incluindo maiúscula, minúscula, número e símbolo.
              </p>
              {passwordStatus ? <p className="text-sm text-slate-300">{passwordStatus}</p> : null}
              <Button type="submit" variant="success" disabled={savingPassword}>
                {savingPassword ? "Alterando..." : "Alterar senha"}
              </Button>
            </form>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
