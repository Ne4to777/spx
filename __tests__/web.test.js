import spx from '../src/modules/web'


describe('tests', () => {
	beforeEach(() => {
		jest.setTimeout(10000)
	})
	it('works with async/await', async () => {
		expect.assertions(1)
		const data = await spx().get()
		expect(data.Title).toEqual('Root HNSC Site Collection')
	})
})
