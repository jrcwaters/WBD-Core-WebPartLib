import { getImportantMail } from '../graphData';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { IMailMessage } from '../types';

interface IRawResponse {
  value: unknown[];
}

/** A fluent MSGraphClientV3 stub: every builder call returns the builder; get() resolves the canned response. */
function mockClient(response: IRawResponse): MSGraphClientV3 {
  const builder: Record<string, unknown> = {};
  const chain = (): unknown => builder;
  builder.api = jest.fn(chain);
  builder.filter = jest.fn(chain);
  builder.select = jest.fn(chain);
  builder.orderby = jest.fn(chain);
  builder.header = jest.fn(chain);
  builder.count = jest.fn(chain);
  builder.top = jest.fn(chain);
  builder.get = jest.fn(() => Promise.resolve(response));
  return builder as unknown as MSGraphClientV3;
}

describe('getImportantMail', () => {
  beforeEach(() => {
    window.sessionStorage.clear();
  });

  it('projects a Graph message onto IMailMessage', async () => {
    const client = mockClient({
      value: [
        {
          id: 'AAA',
          subject: 'Contract sign-off needed',
          from: { emailAddress: { name: 'Jane Doe', address: 'jane@example.com' } },
          receivedDateTime: '2026-07-22T08:30:00Z',
          importance: 'high',
          flag: { flagStatus: 'flagged' },
          isRead: false,
          webLink: 'https://outlook.office.com/mail/id/AAA',
          bodyPreview: 'Please review and sign.'
        }
      ]
    });

    const mail: IMailMessage[] = await getImportantMail(client);

    expect(mail).toHaveLength(1);
    expect(mail[0].id).toBe('AAA');
    expect(mail[0].subject).toBe('Contract sign-off needed');
    expect(mail[0].from).toBe('Jane Doe');
    expect(mail[0].fromAddress).toBe('jane@example.com');
    expect(mail[0].received.toISOString()).toBe('2026-07-22T08:30:00.000Z');
    expect(mail[0].importance).toBe('high');
    expect(mail[0].isFlagged).toBe(true);
    expect(mail[0].isRead).toBe(false);
    expect(mail[0].webLink).toBe('https://outlook.office.com/mail/id/AAA');
    expect(mail[0].preview).toBe('Please review and sign.');
  });

  it('sorts newest first regardless of the order Graph returns', async () => {
    const client = mockClient({
      value: [
        { id: 'OLD', subject: 'Older', receivedDateTime: '2026-07-20T09:00:00Z', importance: 'high' },
        { id: 'NEW', subject: 'Newer', receivedDateTime: '2026-07-22T09:00:00Z', importance: 'high' }
      ]
    });

    const mail = await getImportantMail(client);

    expect(mail.map((m) => m.id)).toEqual(['NEW', 'OLD']);
  });

  it('falls back sensibly when optional fields are missing', async () => {
    const client = mockClient({
      value: [{ id: 'BBB', receivedDateTime: '2026-07-22T10:00:00Z' }]
    });

    const mail = await getImportantMail(client);

    expect(mail[0].subject).toBe('(no subject)');
    expect(mail[0].from).toBe('Unknown sender');
    expect(mail[0].importance).toBe('normal');
    expect(mail[0].isFlagged).toBe(false);
    expect(mail[0].isRead).toBe(true); // absent isRead is treated as read
  });

  it('caches the result and serves rehydrated Dates without a second Graph call', async () => {
    const client = mockClient({
      value: [{ id: 'CCC', subject: 'Cached', receivedDateTime: '2026-07-22T11:00:00Z', importance: 'high' }]
    });
    await getImportantMail(client);
    expect(window.sessionStorage.getItem('hero:mail')).not.toBeNull();

    const client2 = mockClient({ value: [] }); // would blank the result if it were called
    const mail = await getImportantMail(client2);

    expect(client2.get as jest.Mock).not.toHaveBeenCalled();
    expect(mail[0].subject).toBe('Cached');
    expect(mail[0].received instanceof Date).toBe(true);
    expect(mail[0].received.toISOString()).toBe('2026-07-22T11:00:00.000Z');
  });

  it('falls back to an empty array when the Graph call fails', async () => {
    const client = mockClient({ value: [] });
    (client.get as jest.Mock).mockImplementationOnce(() => Promise.reject(new Error('boom')));
    const mail = await getImportantMail(client);
    expect(mail).toEqual([]);
  });
});
