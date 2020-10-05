

export default class Utilities {
  public static createMailFileName(subject: string): string {
    let fileName = subject.replace(/ /g, '-').replace(/:/g, '_').replace(/"/g, '');
    return fileName;
  }
}